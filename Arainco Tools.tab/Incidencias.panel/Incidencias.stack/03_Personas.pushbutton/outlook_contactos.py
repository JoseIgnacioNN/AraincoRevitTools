# -*- coding: utf-8 -*-
"""
Obtiene contactos de Outlook: desde buzón jose.nunez@arainco.cl o desde archivo PST.
Requiere: pip install pywin32
"""
from __future__ import print_function
import os
import re

try:
    unicode
except NameError:
    unicode = str

DOMINIO_FILTRO = "@arainco.cl"
MAILBOX_BUSCAR = "jose.nunez@arainco.cl"
OL_FOLDER_CONTACTS = 10
OL_CONTACT = 40


def _get_emails_from_contact(contact):
    """Extrae emails de un contacto de Outlook."""
    emails = []
    for attr in ("Email1Address", "Email2Address", "Email3Address"):
        try:
            val = getattr(contact, attr, None)
            if val and isinstance(val, (str, unicode)) and val.strip():
                emails.append(val.strip())
        except Exception:
            pass
    if not emails:
        for attr in ("Email1DisplayName", "Email2DisplayName", "Email3DisplayName"):
            try:
                val = getattr(contact, attr, None)
                if val and isinstance(val, (str, unicode)) and val.strip():
                    m = re.search(r'([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})', val)
                    if m:
                        emails.append(m.group(1))
            except Exception:
                pass
    if not emails:
        try:
            pa = contact.PropertyAccessor
            schema = "http://schemas.microsoft.com/mapi/proptag/0x8083001E"
            e = pa.GetProperty(schema)
            if e and str(e).strip():
                emails.append(str(e).strip())
        except Exception:
            pass
    return emails


def _get_nombre_contacto(contact):
    """Obtiene el nombre de un contacto."""
    nombre = getattr(contact, "FullName", None) or ""
    if not nombre or not nombre.strip():
        nombre = (getattr(contact, "FirstName", "") or "") + " " + (getattr(contact, "LastName", "") or "")
        nombre = nombre.strip() or "(Sin nombre)"
    return nombre.strip()


def _iterar_contactos_carpeta(folder, dominio_filtro=None):
    """Itera contactos en una carpeta y subcarpetas."""
    found = []
    try:
        items = folder.Items
        for i in range(1, items.Count + 1):
            try:
                contact = items.Item(i)
                if getattr(contact, "Class", 0) != OL_CONTACT:
                    continue
                nombre = _get_nombre_contacto(contact)
                for email in _get_emails_from_contact(contact):
                    if dominio_filtro and dominio_filtro.lower() not in email.lower():
                        continue
                    found.append({"nombre": nombre or email, "email": email})
                    break
            except Exception:
                continue
        for j in range(1, folder.Folders.Count + 1):
            try:
                sub = folder.Folders.Item(j)
                found.extend(_iterar_contactos_carpeta(sub, dominio_filtro))
            except Exception:
                continue
    except Exception:
        pass
    return found


def obtener_contactos_desde_mailbox(dominio=None):
    """
    Obtiene contactos del buzón jose.nunez@arainco.cl.
    dominio: si se especifica, filtra solo emails que contengan ese dominio (ej: @arainco.cl)
    """
    try:
        import win32com.client
        from win32com.client import constants
    except ImportError:
        return None, "Se requiere pywin32. Ejecuta: pip install pywin32"

    try:
        outlook = win32com.client.gencache.EnsureDispatch("Outlook.Application")
    except Exception as e:
        return None, "No se pudo conectar con Outlook. Asegúrate de tener Outlook instalado y abierto: {}".format(str(e))

    ns = outlook.GetNamespace("MAPI")
    contacts_folder = None
    for i in range(1, ns.Stores.Count + 1):
        store = ns.Stores.Item(i)
        if MAILBOX_BUSCAR.lower() in (store.DisplayName or "").lower():
            try:
                contacts_folder = store.GetDefaultFolder(
                    constants.olFolderContacts if hasattr(constants, "olFolderContacts") else OL_FOLDER_CONTACTS
                )
                break
            except Exception:
                continue
    if not contacts_folder:
        try:
            contacts_folder = ns.GetDefaultFolder(
                constants.olFolderContacts if hasattr(constants, "olFolderContacts") else OL_FOLDER_CONTACTS
            )
        except Exception:
            pass
    if not contacts_folder:
        return None, "No se encontró la carpeta de contactos del buzón {}.".format(MAILBOX_BUSCAR)

    dominio_filtro = dominio or DOMINIO_FILTRO
    resultados = _iterar_contactos_carpeta(contacts_folder, dominio_filtro)
    return resultados, None


def obtener_contactos_desde_pst(pst_path, dominio=None):
    """
    Obtiene contactos de un archivo PST.
    pst_path: ruta absoluta al archivo .pst
    dominio: si se especifica, filtra solo emails que contengan ese dominio
    """
    if not pst_path or not os.path.isfile(pst_path):
        return None, "El archivo PST no existe o la ruta no es válida."

    try:
        import win32com.client
        from win32com.client import constants
    except ImportError:
        return None, "Se requiere pywin32. Ejecuta: pip install pywin32"

    try:
        outlook = win32com.client.gencache.EnsureDispatch("Outlook.Application")
    except Exception as e:
        return None, "No se pudo conectar con Outlook: {}".format(str(e))

    ns = outlook.GetNamespace("MAPI")
    pst_path_abs = os.path.abspath(pst_path)

    try:
        ns.AddStore(pst_path_abs)
    except Exception as e:
        return None, "No se pudo abrir el archivo PST: {}".format(str(e))

    root_pst = None
    pst_norm = os.path.normpath(pst_path_abs).lower()
    try:
        for i in range(1, ns.Stores.Count + 1):
            store = ns.Stores.Item(i)
            if not getattr(store, "IsDataFileStore", False):
                continue
            fp = getattr(store, "FilePath", "") or ""
            if os.path.normpath(fp).lower() == pst_norm:
                root_pst = store.GetRootFolder()
                break
    except Exception:
        pass

    if not root_pst:
        return None, "No se pudo acceder a la raíz del archivo PST."

    def buscar_carpeta_contactos(folder):
        """Busca la carpeta Contactos en el árbol del PST (DefaultItemType=40)."""
        try:
            for j in range(1, folder.Folders.Count + 1):
                sub = folder.Folders.Item(j)
                if getattr(sub, "DefaultItemType", 0) == OL_CONTACT:
                    return sub
                found = buscar_carpeta_contactos(sub)
                if found:
                    return found
        except Exception:
            pass
        return None

    contacts_folder = buscar_carpeta_contactos(root_pst)
    if not contacts_folder:
        try:
            contacts_folder = root_pst.Folders.Item("Contactos")
        except Exception:
            pass
    if not contacts_folder:
        try:
            for j in range(1, root_pst.Folders.Count + 1):
                sub = root_pst.Folders.Item(j)
                if "contact" in (getattr(sub, "Name", "") or "").lower():
                    contacts_folder = sub
                    break
        except Exception:
            pass
    dominio_filtro = dominio or DOMINIO_FILTRO
    if not contacts_folder:
        resultados = _iterar_contactos_carpeta(root_pst, dominio_filtro)
    else:
        resultados = _iterar_contactos_carpeta(contacts_folder, dominio_filtro)

    try:
        ns.RemoveStore(root_pst)
    except Exception:
        pass

    return resultados, None
