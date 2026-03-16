# -*- coding: utf-8 -*-
"""
Recordatorios de incidencias.
Envia recordatorios a las personas asignadas via SMTP.

Uso: python recordatorios_smtp.py
Detener: Ctrl+C
"""

from __future__ import print_function
import os
import sys
import json
import struct
import smtplib
import base64
from datetime import datetime, date
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.utils import formataddr

# Ruta del script para buscar config en el mismo directorio
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH = os.path.join(SCRIPT_DIR, "recordatorios_config.json")

# Paleta Arainco (del logo)
_COLOR_DARK = u"#264A62"   # Azul oscuro - titulos, proyectos
_COLOR_LIGHT = u"#51B2E0"  # Azul claro - acentos
_COLOR_GRAY = u"#B4B4B4"   # Gris - bordes, lineas
_COLOR_WHITE = u"#FFFFFF"


def _log(msg):
    """Imprime con timestamp."""
    from datetime import datetime
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print("[{}] {}".format(ts, msg))
    sys.stdout.flush()


def _deep_merge(base, override):
    """Fusiona override en base recursivamente. override tiene prioridad."""
    result = dict(base)
    for k, v in override.items():
        if k in result and isinstance(result[k], dict) and isinstance(v, dict):
            result[k] = _deep_merge(result[k], v)
        else:
            result[k] = v
    return result


def _load_config():
    """Carga la configuracion desde JSON. Opcionalmente fusiona recordatorios_config.local.json."""
    if not os.path.exists(CONFIG_PATH):
        raise SystemExit(
            "No se encontro recordatorios_config.json en:\n  {}\n"
            "Copia el ejemplo y configura smtp.password.".format(CONFIG_PATH)
        )
    with open(CONFIG_PATH, "r", encoding="utf-8") as f:
        cfg = json.load(f)
    local_path = os.path.join(SCRIPT_DIR, "recordatorios_config.local.json")
    if os.path.isfile(local_path):
        try:
            with open(local_path, "r", encoding="utf-8") as f:
                local_cfg = json.load(f)
            cfg = _deep_merge(cfg, local_cfg)
        except Exception:
            pass
    return cfg


def _parse_fecha_creacion(issue):
    """Extrae la fecha de creacion del issue. Retorna date o None."""
    fecha_str = issue.get("fecha") or issue.get("fecha_creacion") or ""
    if not fecha_str:
        return None
    try:
        if "T" in fecha_str:
            dt = datetime.strptime(fecha_str[:19], "%Y-%m-%dT%H:%M:%S")
        else:
            dt = datetime.strptime(fecha_str[:10], "%Y-%m-%d")
        return dt.date()
    except (ValueError, TypeError):
        return None


def _get_issue_id(issue):
    """Obtiene el ID unico del issue (ej. ISSUE_20260306_221614)."""
    iid = issue.get("id") or ""
    if iid:
        return iid
    path = issue.get("_issue_dir", "")
    return os.path.basename(path) if path else ""


def _load_estado_recordatorios(issues_dir):
    """Carga el archivo de estado de recordatorios enviados."""
    state_path = os.path.join(issues_dir, "_recordatorios.json")
    if not os.path.isfile(state_path):
        return {}
    try:
        with open(state_path, "r", encoding="utf-8-sig") as f:
            return json.load(f)
    except Exception:
        return {}


def _save_estado_recordatorios(issues_dir, state):
    """Guarda el archivo de estado de recordatorios."""
    state_path = os.path.join(issues_dir, "_recordatorios.json")
    try:
        with open(state_path, "w", encoding="utf-8") as f:
            json.dump(state, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


def _debe_enviar_recordatorio(issue, prioridad, cadencia_dias, hoy, estado):
    """
    Decide si se debe enviar recordatorio para este issue.
    cadencia_dias: dias entre recordatorios para esta prioridad
    estado: dict { issue_id: { "ultimo_envio": "YYYY-MM-DD" } }
    - Si hay ultimo_envio: enviar cuando hayan pasado >= cadencia_dias desde ese envio.
    - Si no hay ultimo_envio: enviar cuando hayan pasado >= cadencia_dias desde creacion.
    """
    issue_id = _get_issue_id(issue)
    if not issue_id:
        return False

    fecha_creacion = _parse_fecha_creacion(issue)
    if not fecha_creacion:
        return False

    dias_desde_creacion = (hoy - fecha_creacion).days
    if dias_desde_creacion < 1:
        return False

    ultimo = estado.get(issue_id, {}).get("ultimo_envio")
    if ultimo:
        try:
            fecha_ref = datetime.strptime(ultimo, "%Y-%m-%d").date()
            dias_desde_ref = (hoy - fecha_ref).days
            return dias_desde_ref >= cadencia_dias
        except (ValueError, TypeError):
            pass

    # Primer recordatorio: enviar cuando hayan pasado cadencia_dias desde creacion
    return dias_desde_creacion >= cadencia_dias


def _load_all_issues(issues_dir):
    """
    Carga todas las incidencias del servidor.
    Busca archivos issue.json dentro de carpetas ISSUE_* (por proyecto).
    Usa utf-8-sig para soportar BOM en los JSON.
    """
    issues = []
    if not os.path.isdir(issues_dir):
        return issues
    try:
        for project_name in sorted(os.listdir(issues_dir)):
            project_path = os.path.join(issues_dir, project_name)
            if not os.path.isdir(project_path):
                continue
            try:
                for issue_dir_name in os.listdir(project_path):
                    if not issue_dir_name.startswith("ISSUE_"):
                        continue
                    issue_path = os.path.join(project_path, issue_dir_name)
                    if not os.path.isdir(issue_path):
                        continue
                    json_path = os.path.join(issue_path, "issue.json")
                    if not os.path.exists(json_path):
                        continue
                    try:
                        with open(json_path, "r", encoding="utf-8-sig") as fp:
                            data = json.load(fp)
                    except Exception:
                        continue
                    if not data.get("proyecto_carpeta") and not data.get("proyecto"):
                        data["proyecto_carpeta"] = project_name
                    data["_issue_dir"] = issue_path
                    issues.append(data)
            except OSError:
                continue
    except OSError:
        pass
    return issues


def _filter_issues(issues, estados, prioridades):
    """Filtra incidencias por estado (Abierto, En revision) y prioridad."""
    result = []
    estados_norm = [e.strip().lower() for e in estados if e]
    prioridades_norm = [p.strip().lower() for p in prioridades if p]
    for i in issues:
        est = (i.get("estado") or "").strip().lower()
        pri = (i.get("prioridad") or "").strip().lower()
        if est in estados_norm and pri in prioridades_norm:
            asignado = i.get("asignado_a") or {}
            if isinstance(asignado, dict):
                email = (asignado.get("email") or "").strip()
                if email:
                    result.append(i)
    return result


def _format_issue_line(issue):
    """Formato: Nº de la incidencia - Prioridad - Titulo."""
    n_ord = u"\u00ba"
    numero = issue.get("numero", "?")
    titulo = (issue.get("titulo") or "").strip()
    prioridad = (issue.get("prioridad") or "").strip()
    return u"N{} {} - {} - {}".format(n_ord, numero, prioridad, titulo)


def _escape_html(s):
    """Escapa caracteres especiales para HTML."""
    return (s or u"").replace(u"&", u"&amp;").replace(u"<", u"&lt;").replace(u">", u"&gt;").replace(u'"', u"&quot;")


def _get_image_dimensions(path):
    """Obtiene ancho y alto de imagen PNG o JPEG. Retorna (width, height) o (None, None)."""
    try:
        ext = os.path.splitext(path)[1].lower()
        with open(path, "rb") as f:
            if ext == ".png":
                header = f.read(24)
                if header[:8] == b"\x89PNG\r\n\x1a\n" and header[12:16] == b"IHDR":
                    w, h = struct.unpack(">II", header[16:24])
                    return w, h
            elif ext in (".jpg", ".jpeg"):
                if f.read(2) != b"\xff\xd8":
                    return None, None
                while True:
                    chunk = f.read(4)
                    if len(chunk) < 4:
                        break
                    marker, size = struct.unpack(">HH", chunk)
                    if marker in (0xFFC0, 0xFFC1, 0xFFC2):
                        data = f.read(5)
                        if len(data) >= 5:
                            h, w = struct.unpack(">HH", data[1:5])
                            return w, h
                        break
                    f.seek(size - 2, 1)
    except Exception:
        pass
    return None, None


def _load_logo_base64(logo_path, logo_cfg=None):
    """
    Carga el logo como data URI base64 y sus dimensiones.
    Retorna (data_uri, display_width, display_height) para escalar a max 50px de alto.
    Si no existe: (u"", None, None). logo_cfg puede tener ancho/alto para forzar dimensiones.
    """
    if not logo_path:
        return u"", None, None
    path = os.path.join(SCRIPT_DIR, logo_path) if not os.path.isabs(logo_path) else logo_path
    if not os.path.isfile(path):
        return u"", None, None
    logo_cfg = logo_cfg or {}
    cfg_w = logo_cfg.get("ancho")
    cfg_h = logo_cfg.get("alto")
    if cfg_w is not None and cfg_h is not None and cfg_w > 0 and cfg_h > 0:
        try:
            cfg_w, cfg_h = int(cfg_w), int(cfg_h)
        except (TypeError, ValueError):
            cfg_w, cfg_h = None, None
    try:
        with open(path, "rb") as f:
            data = base64.b64encode(f.read()).decode("ascii")
        ext = os.path.splitext(path)[1].lower()
        mime = "image/png" if ext == ".png" else "image/jpeg" if ext in (".jpg", ".jpeg") else "image/png"
        data_uri = u"data:{};base64,{}".format(mime, data)

        if cfg_w and cfg_h:
            return data_uri, cfg_w, cfg_h
        orig_w, orig_h = _get_image_dimensions(path)
        if orig_w and orig_h and orig_h > 0:
            display_h = 50
            display_w = int(orig_w * display_h / orig_h)
            if display_w > 300:
                display_w = 300
                display_h = int(orig_h * display_w / orig_w)
            return data_uri, display_w, display_h
        return data_uri, 180, 50
    except Exception:
        return u"", None, None


def _build_email_body_agrupado(issues_by_project, person_name, remitente_nombre, logo_data_uri=u"", logo_w=None, logo_h=None):
    """
    Construye el cuerpo del correo con incidencias agrupadas por proyecto.
    Usa MJML para compatibilidad Outlook clasico y nuevo. Retorna (plain_text, html).
    """
    lines = [
        u"Estimado/a {},".format(person_name or "colega"),
        u"",
        u"Tienes las siguientes incidencias sin resolver aun:",
        u"",
    ]

    for proyecto in sorted(issues_by_project.keys()):
        issues = issues_by_project[proyecto]
        lines.append(proyecto)
        for issue in issues:
            lines.append(_format_issue_line(issue))
        lines.append(u"")

    lines.extend([
        u"Por favor revisa el servidor de incidencias para mas detalles.",
        u"",
        u"Saludos,",
        u"{}".format(remitente_nombre),
    ])

    plain = u"\r\n".join(lines)

    mjml_str = _build_mjml_recordatorios(
        issues_by_project, person_name, remitente_nombre, logo_data_uri, logo_w, logo_h
    )
    html = _mjml_to_html(mjml_str)

    if not html:
        _log("  Advertencia: mjml-python no disponible (pip install mjml-python), usando HTML basico")
        html = _build_email_body_html_fallback(
            issues_by_project, person_name, remitente_nombre, logo_data_uri, logo_w, logo_h
        )

    return plain, html


def _build_mjml_recordatorios(issues_by_project, person_name, remitente_nombre, logo_data_uri=u"", logo_w=None, logo_h=None):
    """Construye el cuerpo del correo en MJML."""
    w = logo_w if logo_w else 180
    h = logo_h if logo_h else 50
    person_esc = _escape_html(person_name or "colega")
    rem_esc = _escape_html(remitente_nombre)

    parts = [
        u'<mjml>',
        u'<mj-body width="600px" background-color="{}">'.format(_COLOR_WHITE),
    ]

    if logo_data_uri:
        parts.append(u'<mj-section background-color="{}" padding="20px 24px">'.format(_COLOR_WHITE))
        parts.append(u'<mj-column><mj-image src="{}" width="{}px" height="{}px" alt="Arainco" /></mj-column>'.format(logo_data_uri, w, h))
    else:
        parts.append(u'<mj-section background-color="{}" padding="16px 24px">'.format(_COLOR_DARK))
        parts.append(u'<mj-column><mj-text color="{}" font-size="20px" font-weight="bold">ARAINCO</mj-text><mj-text color="{}" font-size="12px">INGENIERIA ESTRUCTURAL</mj-text></mj-column>'.format(_COLOR_WHITE, _COLOR_LIGHT))
    parts.append(u'</mj-section>')

    parts.append(u'<mj-section background-color="{}" padding="24px">'.format(_COLOR_WHITE))
    parts.append(u'<mj-column><mj-text color="{}" font-size="14px">Estimado/a {},</mj-text><mj-text color="#555555" font-size="14px">Tienes las siguientes incidencias sin resolver aun:</mj-text></mj-column>'.format(_COLOR_DARK, person_esc))
    parts.append(u'</mj-section>')

    for proyecto in sorted(issues_by_project.keys()):
        issues = issues_by_project[proyecto]
        proy_esc = _escape_html(proyecto)
        issues_lines = u"<br/>".join(_escape_html(_format_issue_line(i)) for i in issues)
        parts.append(u'<mj-section background-color="#f8fafc" padding="12px 16px">')
        parts.append(u'<mj-column><mj-text color="{}" font-weight="bold" font-size="13px">{}</mj-text><mj-text color="#444444" font-size="13px">{}</mj-text></mj-column>'.format(_COLOR_DARK, proy_esc, issues_lines))
        parts.append(u'</mj-section>')

    parts.append(u'<mj-section background-color="{}" padding="24px">'.format(_COLOR_WHITE))
    parts.append(u'<mj-column><mj-text color="#666666" font-size="13px">Por favor revisa el servidor de incidencias para mas detalles.</mj-text><mj-text color="{}" font-size="14px">Saludos,<br/><strong>{}</strong></mj-text></mj-column>'.format(_COLOR_DARK, rem_esc))
    parts.append(u'</mj-section>')

    parts.append(u'<mj-section background-color="{}" padding="12px 24px">'.format(_COLOR_DARK))
    parts.append(u'<mj-column><mj-text color="{}" font-size="11px">Arainco Ingenieria Estructural - Recordatorio de incidencias</mj-text></mj-column>'.format(_COLOR_WHITE))
    parts.append(u'</mj-section>')
    parts.append(u'</mj-body></mjml>')
    return u"".join(parts)


def _mjml_to_html(mjml_str):
    """Compila MJML a HTML. Retorna None si falla."""
    try:
        from mjml import mjml2html
        result = mjml2html(mjml_str)
        return result if isinstance(result, str) else getattr(result, "html", str(result))
    except (ImportError, Exception):
        return None


def _build_email_body_html_fallback(issues_by_project, person_name, remitente_nombre, logo_data_uri=u"", logo_w=None, logo_h=None):
    """Fallback HTML si MJML no esta disponible."""
    w = logo_w if logo_w else 180
    h = logo_h if logo_h else 50
    parts = [u"<div style='font-family:Segoe UI,Arial,sans-serif;font-size:14px;color:#333;max-width:600px'>"]
    if logo_data_uri:
        parts.append(u"<div style='background:{};padding:20px 24px;border-bottom:3px solid {}'><img src='{}' alt='Arainco' width='{}' height='{}' style='display:block;border:0' /></div>".format(_COLOR_WHITE, _COLOR_LIGHT, logo_data_uri, w, h))
    else:
        parts.append(u"<div style='background:{};padding:16px 24px;border-bottom:3px solid {}'><span style='color:{};font-size:20px;font-weight:bold'>ARAINCO</span> <span style='color:{};font-size:12px'>INGENIERIA ESTRUCTURAL</span></div>".format(_COLOR_DARK, _COLOR_LIGHT, _COLOR_WHITE, _COLOR_LIGHT))
    parts.append(u"<div style='padding:24px;background:{}'><p style='color:{}'>Estimado/a {},</p><p style='color:#555'>Tienes las siguientes incidencias sin resolver aun:</p>".format(_COLOR_WHITE, _COLOR_DARK, _escape_html(person_name or "colega")))
    for proyecto in sorted(issues_by_project.keys()):
        issues = issues_by_project[proyecto]
        parts.append(u"<div style='margin-bottom:16px;padding:12px 16px;background:#f8fafc;border-left:4px solid {}'><div style='color:{};font-weight:bold;margin-bottom:8px'>{}</div>".format(_COLOR_LIGHT, _COLOR_DARK, _escape_html(proyecto)))
        for issue in issues:
            parts.append(u"<div style='color:#444;padding:4px 0;border-bottom:1px solid {}'>{}</div>".format(_COLOR_GRAY, _escape_html(_format_issue_line(issue))))
        parts.append(u"</div>")
    parts.append(u"<p style='color:#666;margin-top:24px'>Por favor revisa el servidor de incidencias para mas detalles.</p><p style='color:{}'>Saludos,<br/><strong>{}</strong></p></div>".format(_COLOR_DARK, _escape_html(remitente_nombre)))
    parts.append(u"<div style='padding:12px 24px;background:{};color:{};font-size:11px;border-top:1px solid {}'>Arainco Ingenieria Estructural - Recordatorio de incidencias</div></div>".format(_COLOR_DARK, _COLOR_WHITE, _COLOR_LIGHT))
    return u"".join(parts)


def _send_smtp(host, port, use_tls, use_ssl, user, password, from_email, from_name, to_email, subject, body_plain, body_html):
    """Envia un correo via SMTP (multipart: plain + HTML). Port 465 usa SSL; 587 usa STARTTLS."""
    msg = MIMEMultipart("alternative")
    msg["Subject"] = subject
    msg["From"] = formataddr((from_name, from_email))
    msg["To"] = to_email
    msg.attach(MIMEText(body_plain, "plain", "utf-8"))
    msg.attach(MIMEText(body_html, "html", "utf-8"))

    if use_ssl or port == 465:
        conn = smtplib.SMTP_SSL(host, port)
    else:
        conn = smtplib.SMTP(host, port)
        if use_tls:
            conn.starttls()
    try:
        conn.login(user, password)
        conn.sendmail(from_email, [to_email], msg.as_string())
    finally:
        conn.quit()


def main(modo_prueba=False):
    cfg = _load_config()
    smtp_cfg = cfg.get("smtp", {})
    rem_cfg = cfg.get("remitente", {})
    inc_cfg = cfg.get("incidencias", {})

    issues_dir = inc_cfg.get("ruta_servidor", "Y:\\00_SERVIDOR DE INCIDENCIAS")
    estados = inc_cfg.get("estados", ["Abierto", "En revision"])
    cadencia_dias_raw = inc_cfg.get("cadencia_dias", {"Critica": 1, "Alta": 2, "Media": 4, "Baja": 6})
    cadencia_dias = {}
    for pri, dias in cadencia_dias_raw.items():
        if isinstance(pri, str) and isinstance(dias, (int, float)):
            cadencia_dias[pri.strip()] = int(dias)
    solo_laborables = inc_cfg.get("solo_laborables", True)

    from_email = rem_cfg.get("email", "")
    from_name = rem_cfg.get("nombre", "Arainco Notificaciones")

    if not os.path.isdir(issues_dir):
        _log("ERROR: Ruta de incidencias no accesible: {}".format(issues_dir))
        sys.exit(1)

    if not cadencia_dias:
        _log("ERROR: cadencia_dias no puede estar vacia en config")
        sys.exit(1)

    logo_cfg = cfg.get("logo", {})
    logo_path = logo_cfg.get("ruta", "logo.png")
    logo_data_uri, logo_w, logo_h = _load_logo_base64(logo_path, logo_cfg)
    if not logo_data_uri and logo_path:
        _log("  Nota: Logo no encontrado ({}) - se usa encabezado alternativo".format(logo_path))

    smtp_server = smtp_cfg.get("server", "smtp.gmail.com")
    smtp_port = int(smtp_cfg.get("port", 587))
    smtp_tls = smtp_cfg.get("use_tls", True)
    smtp_ssl = smtp_cfg.get("use_ssl", smtp_port == 465)
    smtp_user = smtp_cfg.get("user", "")
    smtp_pass = smtp_cfg.get("password", "")
    if not smtp_pass or smtp_pass == "REEMPLAZAR_CON_PASSWORD_O_APP_PASSWORD":
        smtp_pass = os.environ.get("RECORDATORIOS_SMTP_PASSWORD", "")

    if not smtp_user or not smtp_pass:
        _log("ERROR: Configura smtp.user y smtp.password en recordatorios_config.json")
        _log("      O define la variable de entorno RECORDATORIOS_SMTP_PASSWORD")
        sys.exit(1)

    from_email = from_email or smtp_user

    def _send_smtp_wrapper(to, subj, plain, html):
        _send_smtp(
            smtp_server, smtp_port, smtp_tls, smtp_ssl,
            smtp_user, smtp_pass,
            from_email, from_name, to, subj, plain, html
        )

    send_fn = _send_smtp_wrapper
    modo_label = "SMTP ({}:{})".format(smtp_server, smtp_port)

    hoy = date.today()
    if solo_laborables and hoy.weekday() >= 5:
        _log("Hoy es fin de semana. No se envian recordatorios (solo_laborables=true).")
        return

    n_ord = u"\u00ba"
    _log("Iniciando recordatorios (cadencia por dias)" + (" [MODO PRUEBA]" if modo_prueba else ""))
    _log("  Modo: {}".format(modo_label))
    _log("  Remitente: {} <{}>".format(from_name, from_email))
    _log("  Incidencias: {}".format(issues_dir))
    _log("  Estados: {}".format(estados))
    _log("  Cadencia (dias): {}".format(cadencia_dias))
    _log("  Fecha: {}".format(hoy.isoformat()))
    _log("")

    issues = _load_all_issues(issues_dir)
    estado = _load_estado_recordatorios(issues_dir)
    estado_modified = False

    for prioridad, dias_cadencia in cadencia_dias.items():
        filtered = _filter_issues(issues, estados, [prioridad])
        if modo_prueba:
            to_send = filtered
        else:
            to_send = [i for i in filtered if _debe_enviar_recordatorio(i, prioridad, dias_cadencia, hoy, estado)]

        if not to_send:
            _log("  [{}] Sin incidencias que requieran recordatorio hoy".format(prioridad))
            continue

        by_person = {}
        for issue in to_send:
            asignado = issue.get("asignado_a") or {}
            if not isinstance(asignado, dict):
                continue
            to_email = (asignado.get("email") or "").strip()
            if not to_email:
                continue
            if to_email not in by_person:
                by_person[to_email] = {"nombre": asignado.get("nombre", ""), "issues": []}
            by_person[to_email]["issues"].append(issue)

        _log("  [{}] {} persona(s), {} incidencia(s) a recordar".format(prioridad, len(by_person), len(to_send)))
        for to_email, data in by_person.items():
            person_issues = data["issues"]
            person_name = data["nombre"]

            by_project = {}
            for issue in person_issues:
                proy = (issue.get("proyecto_carpeta") or issue.get("proyecto") or "").strip() or "Sin proyecto"
                if proy not in by_project:
                    by_project[proy] = []
                by_project[proy].append(issue)

            total = len(person_issues)
            subject = u"Recordatorio [{}]: {} incidencia(s) pendiente(s)".format(prioridad, total)
            body_plain, body_html = _build_email_body_agrupado(
                by_project, person_name, from_name, logo_data_uri, logo_w, logo_h
            )
            try:
                send_fn(to_email, subject, body_plain, body_html)
                _log("    Enviado a {} ({} incidencia(s))".format(to_email, total))
                if not modo_prueba:
                    for issue in person_issues:
                        issue_id = _get_issue_id(issue)
                        if issue_id:
                            if issue_id not in estado:
                                estado[issue_id] = {}
                            estado[issue_id]["ultimo_envio"] = hoy.isoformat()
                            estado_modified = True
            except Exception as e:
                _log("    ERROR al enviar a {}: {}".format(to_email, e))
                _log("      Incidencias afectadas: {}".format(
                    ", ".join(_get_issue_id(i) for i in person_issues)))
                _log("      Se reintentara en la proxima ejecucion.")

    if estado_modified:
        _save_estado_recordatorios(issues_dir, estado)

    _log("")
    _log("Proceso completado.")


if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description="Recordatorios de incidencias")
    parser.add_argument("--prueba", action="store_true", help="Modo prueba: ignora cadencia y envia a todas las incidencias pendientes")
    args = parser.parse_args()
    main(modo_prueba=args.prueba)
