# -*- coding: utf-8 -*-
"""
Extrae contactos de un archivo PST a JSON (sin Outlook).
Requiere: pip install libpff-python-windows  (Python 3.10+)

Uso: python pst_to_json.py archivo.pst [salida.json]
"""
from __future__ import print_function
import sys
import os
import io
import json
import re

def extract_contacts_from_pst(pst_path, output_path=None):
    """Extrae contactos del PST usando pypff y guarda en JSON."""
    try:
        import pypff
    except ImportError:
        print("ERROR: Instala pypff con: pip install libpff-python-windows")
        print("Requiere Python 3.10 o superior.")
        return False

    if not output_path:
        base = os.path.splitext(pst_path)[0]
        output_path = base + "_contactos.json"

    contactos = []
    try:
        pst = pypff.file()
        pst.open(pst_path)
        root = pst.get_root_folder()
        _extract_from_folder(root, contactos, pypff)
        pst.close()
    except Exception as e:
        print("Error al abrir PST: {}".format(e))
        return False

    try:
        with io.open(output_path, "w", encoding="utf-8") as f:
            json.dump(contactos, f, ensure_ascii=False, indent=2)
        print("Guardados {} contactos en {}".format(len(contactos), output_path))
        return True
    except Exception as e:
        print("Error al guardar JSON: {}".format(e))
        return False


def _extract_from_folder(folder, contactos, pypff_mod):
    """Recorre carpetas e items buscando contactos."""
    try:
        for item in folder.items:
            try:
                if isinstance(item, pypff_mod.folder):
                    _extract_from_folder(item, contactos, pypff_mod)
                else:
                    c = _item_to_contact(item)
                    if c:
                        contactos.append(c)
            except Exception:
                pass
    except Exception:
        pass


def _item_to_contact(item):
    """Intenta extraer nombre y email de un item (contacto, mensaje, etc.)."""
    nombre = ""
    email = ""
    try:
        # Display name (contactos y otros items)
        if hasattr(item, "get_display_name"):
            nombre = (item.get_display_name() or "").strip()
        if hasattr(item, "get_utf8_display_name"):
            n = item.get_utf8_display_name()
            if n and not nombre:
                nombre = (n.decode("utf-8", errors="replace") if isinstance(n, bytes) else str(n)).strip()
        # Subject (mensajes)
        if not nombre and hasattr(item, "get_subject"):
            nombre = (item.get_subject() or "").strip()
        # Record set / MAPI properties (contactos) - si la API existe
        try:
            if hasattr(item, "get_record_set"):
                rs = item.get_record_set()
                if rs and hasattr(rs, "get_number_of_entries"):
                    for i in range(rs.get_number_of_entries()):
                        try:
                            entry = rs.get_entry_by_index(i)
                            if entry and hasattr(entry, "get_data") and entry.get_data():
                                data = entry.get_data()
                                if isinstance(data, bytes) and len(data) > 0:
                                    s = data.decode("utf-8", errors="replace").strip()
                                    if "@" in s and "." in s:
                                        m = re.search(r'[\w.-]+@[\w.-]+\.\w+', s)
                                        if m and not email:
                                            email = m.group(0)
                                    elif s and not nombre and len(s) < 100:
                                        nombre = s
                        except Exception:
                            pass
        except Exception:
            pass
        # Cuerpo y headers (emails)
        if hasattr(item, "get_plain_text_body"):
            body = (item.get_plain_text_body() or "") or ""
            m = re.search(r'[\w.-]+@[\w.-]+\.\w+', body)
            if m and not email:
                email = m.group(0)
        if hasattr(item, "get_transport_headers"):
            headers = (item.get_transport_headers() or "") or ""
            m = re.search(r'[\w.-]+@[\w.-]+\.\w+', headers)
            if m and not email:
                email = m.group(0)
        if hasattr(item, "get_html_body"):
            html = (item.get_html_body() or "") or ""
            m = re.search(r'[\w.-]+@[\w.-]+\.\w+', html)
            if m and not email:
                email = m.group(0)
        if nombre or email:
            return {"nombre": nombre or "(Sin nombre)", "email": email or ""}
    except Exception:
        pass
    return None


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Uso: python pst_to_json.py archivo.pst [salida.json]")
        sys.exit(1)
    pst_path = sys.argv[1]
    out_path = sys.argv[2] if len(sys.argv) > 2 else None
    if not os.path.isfile(pst_path):
        print("No existe: {}".format(pst_path))
        sys.exit(1)
    ok = extract_contacts_from_pst(pst_path, out_path)
    sys.exit(0 if ok else 1)
