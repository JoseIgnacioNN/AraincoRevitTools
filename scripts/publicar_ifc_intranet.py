# -*- coding: utf-8 -*-
"""Exporta el RVT abierto a IFC y lo publica en la intranet (/modelos)."""

from __future__ import print_function

import codecs
import json
import os
import re
import shutil
import tempfile
from datetime import datetime

from Autodesk.Revit.DB import (
    FilteredElementCollector,
    IFCExportOptions,
    IFCVersion,
    Transaction,
    TransactionStatus,
    View3D,
)
from Autodesk.Revit.UI import TaskDialog

_DIALOG = u"Arainco: Publicar IFC intranet"
_SLUG_RE = re.compile(r"^[A-Za-z0-9._-]{1,80}$")
_INTRANET_URLS = (
    u"http://127.0.0.1:3000",
    u"http://localhost:3000",
)


def _u(value):
    try:
        return unicode(value)
    except NameError:
        return str(value)
    except Exception:
        return u""


def _show(uiapp, instruction, content=u""):
    try:
        from bimtools_instruction_dialog import show_message_dialog
        from revit_wpf_window_position import revit_main_hwnd

        hwnd = revit_main_hwnd(uiapp) if uiapp else None
        show_message_dialog(
            _DIALOG,
            instruction=instruction,
            content=content or u"",
            ok_text=u"Entendido",
            hwnd_revit=hwnd,
            uiapp=uiapp,
        )
        return
    except Exception:
        pass
    msg = instruction
    if content:
        msg = instruction + u"\n\n" + content
    TaskDialog.Show(_DIALOG, msg)


def _find_intranet_root(start_dir):
    cursor = os.path.abspath(start_dir)
    for _ in range(16):
        here = cursor
        if os.path.isfile(os.path.join(here, u"package.json")) and os.path.isdir(
            os.path.join(here, u"src")
        ):
            if os.path.basename(here) == u"intranet-arainco":
                return here
        sibling = os.path.join(os.path.dirname(cursor), u"intranet-arainco")
        if os.path.isfile(os.path.join(sibling, u"package.json")):
            return sibling
        parent = os.path.dirname(cursor)
        if parent == cursor:
            break
        cursor = parent
    return None


def _read_env(intranet):
    env_path = os.path.join(intranet, u".env.local")
    out = {}
    if not os.path.isfile(env_path):
        return out
    try:
        with codecs.open(env_path, u"r", u"utf-8") as fh:
            for line in fh:
                line = line.strip()
                if not line or line.startswith(u"#") or u"=" not in line:
                    continue
                key, val = line.split(u"=", 1)
                out[key.strip()] = val.strip().strip(u'"').strip(u"'")
    except Exception:
        return out
    return out


def _intranet_bases(env):
    urls = []
    for key in (u"INTRANET_URL", u"NEXT_PUBLIC_SITE_URL"):
        val = (env.get(key) or u"").rstrip(u"/")
        if val and val not in urls:
            urls.append(val)
    for extra in _INTRANET_URLS:
        if extra not in urls:
            urls.append(extra)
    return urls


def _slug_from_doc(doc):
    path = doc.PathName
    if path:
        base = os.path.splitext(os.path.basename(_u(path)))[0]
    else:
        base = _u(doc.Title) or u"modelo"
    clean = re.sub(r"[^A-Za-z0-9._-]+", u"_", base)[:80]
    if _SLUG_RE.match(clean):
        return clean
    return u"modelo"


def _pick_3d_view(doc, uidoc):
    active = uidoc.ActiveView if uidoc else None
    if isinstance(active, View3D) and not active.IsTemplate:
        return active
    for view in FilteredElementCollector(doc).OfClass(View3D):
        if not view.IsTemplate:
            return view
    return None


def _ifc_version():
    for name in (u"IFC4RV", u"IFC4", u"IFC2x3CV2"):
        if hasattr(IFCVersion, name):
            return getattr(IFCVersion, name)
    return IFCVersion.Default


def _write_catalog(catalog_path, entry):
    rows = []
    if os.path.isfile(catalog_path):
        try:
            with codecs.open(catalog_path, u"r", u"utf-8") as fh:
                raw = json.loads(fh.read() or u"[]")
            if isinstance(raw, list):
                rows = [r for r in raw if r.get(u"slug") != entry[u"slug"]]
        except Exception:
            rows = []
    rows.append(entry)
    payload = json.dumps(rows, indent=2, ensure_ascii=False)
    with codecs.open(catalog_path, u"w", u"utf-8") as fh:
        fh.write(payload)


def _open_browser(url):
    try:
        os.startfile(url)
        return True
    except Exception:
        pass
    try:
        from System.Diagnostics import Process

        Process.Start(url)
        return True
    except Exception:
        return False


def _http_json(url, payload, token):
    from System.IO import StreamReader
    from System.Net import WebClient, WebException
    from System.Text import Encoding

    client = WebClient()
    client.Encoding = Encoding.UTF8
    client.Headers.Add(u"Content-Type", u"application/json")
    client.Headers.Add(u"Accept", u"application/json")
    if token:
        client.Headers.Add(u"x-publish-token", token)
    try:
        raw = client.UploadString(url, u"POST", json.dumps(payload))
        return json.loads(raw) if raw else {}
    except WebException as ex:
        body = u""
        try:
            if ex.Response is not None:
                reader = StreamReader(ex.Response.GetResponseStream())
                body = reader.ReadToEnd()
                reader.Close()
        except Exception:
            body = u""
        raise Exception(u"{0} {1}".format(_u(ex.Message), body[:400]))
    finally:
        client.Dispose()


def _put_ifc(signed_url, bearer, ifc_path):
    from System.IO import File
    from System.Net import WebClient

    client = WebClient()
    client.Headers.Add(u"Authorization", u"Bearer " + bearer)
    client.Headers.Add(u"x-upsert", u"true")
    client.Headers.Add(u"Content-Type", u"application/octet-stream")
    try:
        data = File.ReadAllBytes(ifc_path)
        client.UploadData(signed_url, u"PUT", data)
    finally:
        client.Dispose()


def _publish_remote(base, token, ifc_path, slug, title):
    if not token:
        raise Exception(u"Falta INTRANET_PUBLISH_TOKEN en .env.local")
    signed = _http_json(
        base.rstrip(u"/") + u"/api/models/upload-url",
        {u"slug": slug, u"filename": slug + u".ifc"},
        token,
    )
    if not signed.get(u"signedUrl") or not signed.get(u"token"):
        raise Exception(u"La intranet no devolvió URL de subida.")
    _put_ifc(signed[u"signedUrl"], signed[u"token"], ifc_path)
    _http_json(
        base.rstrip(u"/") + u"/api/models",
        {
            u"slug": slug,
            u"title": title,
            u"code": slug,
            u"source": u"Publicado desde Revit",
            u"file": slug + u".ifc",
        },
        token,
    )
    return True


def _first_alive_base(bases, token):
    from System.Net import WebException, WebRequest

    for base in bases:
        try:
            req = WebRequest.Create(base.rstrip(u"/") + u"/api/models")
            req.Method = u"GET"
            req.Timeout = 4000
            if token:
                req.Headers.Add(u"x-publish-token", token)
            resp = req.GetResponse()
            resp.Close()
            return base
        except WebException as ex:
            if ex.Response is not None:
                try:
                    ex.Response.Close()
                except Exception:
                    pass
                return base
        except Exception:
            continue
    return u""


def _rollback_if_open(txn):
    try:
        if txn.GetStatus() == TransactionStatus.Started:
            txn.RollBack()
    except Exception:
        pass


def _export_ifc(doc, uidoc, slug, view3d):
    if view3d is not None:
        try:
            uidoc.ActiveView = view3d
        except Exception:
            pass

    temp_dir = tempfile.mkdtemp(prefix=u"arainco_ifc_")
    options = IFCExportOptions()
    options.FileVersion = _ifc_version()
    options.WallAndColumnSplitting = False
    options.ExportBaseQuantities = False
    try:
        options.AddOption(u"SitePlacement", u"2")
    except Exception:
        pass
    if view3d is not None:
        try:
            options.FilterViewId = view3d.Id
        except Exception:
            pass
        try:
            options.AddOption(u"VisibleElementsOfCurrentView", u"true")
        except Exception:
            pass

    # Document.Export(IFC) exige transacción abierta: el exporter escribe
    # GUIDs IFC en el modelo. Rollback para no ensuciar el RVT.
    txn = Transaction(doc, u"Arainco: Publicar IFC intranet")
    exported = False
    try:
        txn.Start()
        try:
            fail = txn.GetFailureHandlingOptions()
            fail.SetClearAfterRollback(True)
            txn.SetFailureHandlingOptions(fail)
        except Exception:
            pass
        exported = doc.Export(temp_dir, slug, options)
    finally:
        _rollback_if_open(txn)

    produced = os.path.join(temp_dir, slug + u".ifc")
    if not exported or not os.path.isfile(produced):
        candidates = [
            os.path.join(temp_dir, n)
            for n in os.listdir(temp_dir)
            if n.lower().endswith(u".ifc")
        ]
        produced = candidates[0] if candidates else None
    return temp_dir, produced


def run(uiapp):
    uidoc = uiapp.ActiveUIDocument
    if uidoc is None:
        _show(uiapp, u"Abre un modelo antes de publicar.")
        return
    doc = uidoc.Document
    if doc.IsFamilyDocument:
        _show(uiapp, u"Abre un proyecto (.rvt), no una familia.")
        return

    intranet = _find_intranet_root(os.path.abspath(__file__))
    if not intranet:
        _show(
            uiapp,
            u"No se encontró la carpeta intranet-arainco.",
            u"Debe estar junto a BIMTools.extension en CustomRevitExtensions.",
        )
        return

    slug = _slug_from_doc(doc)
    title = _u(doc.Title) or slug
    view3d = _pick_3d_view(doc, uidoc)

    dest_ifc = None
    temp_dir = None
    try:
        temp_dir, produced = _export_ifc(doc, uidoc, slug, view3d)
        if not produced:
            _show(
                uiapp,
                u"Revit no generó el IFC.",
                u"Comprueba que el exporter IFC de Autodesk esté instalado.",
            )
            return

        files_dir = os.path.join(intranet, u"data", u"models", u"files")
        if not os.path.isdir(files_dir):
            os.makedirs(files_dir)
        dest_ifc = os.path.join(files_dir, slug + u".ifc")
        shutil.copyfile(produced, dest_ifc)

        entry = {
            u"slug": slug,
            u"title": title,
            u"code": slug,
            u"source": u"Publicado desde Revit",
            u"file": slug + u".ifc",
            u"publishedAt": datetime.now().strftime(u"%Y-%m-%d"),
        }
        _write_catalog(os.path.join(intranet, u"data", u"models", u"catalog.json"), entry)
    except Exception as ex:
        _show(uiapp, u"Error al exportar el IFC.", _u(ex))
        return
    finally:
        if temp_dir and os.path.isdir(temp_dir):
            try:
                shutil.rmtree(temp_dir)
            except Exception:
                pass

    if not dest_ifc:
        return

    env = _read_env(intranet)
    token = env.get(u"INTRANET_PUBLISH_TOKEN") or u""
    bases = _intranet_bases(env)
    remote_err = u""
    remote_ok = False
    prod_bases = [b for b in bases if u"localhost" not in b and u"127.0.0.1" not in b]
    for base in prod_bases:
        try:
            _publish_remote(base, token, dest_ifc, slug, title)
            remote_ok = True
            view_url = base.rstrip(u"/") + u"/modelos/" + slug
            _open_browser(view_url)
            extra = (
                u"El modelo quedó en la intranet del equipo.\n"
                u"Cualquier correo @arainco.cl puede abrirlo:\n{0}".format(view_url)
            )
            break
        except Exception as ex:
            remote_err = _u(ex)
            continue

    if not remote_ok:
        alive = _first_alive_base(bases, token)
        view_url = (alive or u"http://localhost:3000").rstrip(u"/") + u"/modelos/" + slug
        if alive:
            _open_browser(view_url)
            extra = u"El visor local se abrió. Si no ves el modelo, recarga la página."
            if prod_bases:
                extra += (
                    u"\n\nNo se pudo publicar al equipo ({0}). "
                    u"Revisa INTRANET_URL e INTRANET_PUBLISH_TOKEN."
                ).format(remote_err[:280] if remote_err else u"sin URL de producción")
        else:
            extra = (
                u"El IFC quedó en este PC. Arranca npm run dev o configura "
                u"INTRANET_URL para el equipo.\n{0}"
            ).format(view_url)
            if remote_err:
                extra += u"\n\n" + remote_err[:400]

    view_hint = _u(view3d.Name) if view3d is not None else u"documento completo"
    _show(
        uiapp,
        u"IFC publicado: {0}".format(slug),
        u"Vista exportada: {0}\nArchivo: {1}\n\n{2}".format(view_hint, dest_ifc, extra),
    )
