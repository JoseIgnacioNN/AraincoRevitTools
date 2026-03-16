# -*- coding: utf-8 -*-
"""
Importar contactos de Outlook al directorio de Personas.
Filtra solo los contactos cuya dirección de correo contiene @arainco.cl
Requiere: pip install pywin32
"""
from __future__ import print_function
import os

try:
    unicode
except NameError:
    unicode = str
import json
import sys

ISSUES_DIR = u"Y:\\00_SERVIDOR DE INCIDENCIAS"
PERSONAS_FILE = os.path.join(ISSUES_DIR, "personas.json")
DOMINIO_FILTRO = "@arainco.cl"
MAILBOX_BUSCAR = "jose.nunez@arainco.cl"  # Buzón donde buscar contactos


def obtener_emails_arainco_desde_outlook():
    """Obtiene contactos de Outlook cuyo email contiene @arainco.cl"""
    try:
        import win32com.client
        from win32com.client import constants
    except ImportError:
        print("ERROR: Se requiere pywin32. Ejecuta: pip install pywin32")
        sys.exit(1)

    try:
        outlook = win32com.client.gencache.EnsureDispatch("Outlook.Application")
    except Exception as e:
        print("ERROR: No se pudo conectar con Outlook. Asegurate de tener Outlook instalado y abierto.")
        print(str(e))
        sys.exit(1)

    ns = outlook.GetNamespace("MAPI")
    contacts_folder = None
    for i in range(1, ns.Stores.Count + 1):
        store = ns.Stores.Item(i)
        if MAILBOX_BUSCAR.lower() in (store.DisplayName or "").lower():
            try:
                contacts_folder = store.GetDefaultFolder(constants.olFolderContacts if hasattr(constants, "olFolderContacts") else 10)
                print("Buscando en buzón: {}".format(store.DisplayName))
                break
            except Exception:
                continue
    if not contacts_folder:
        try:
            contacts_folder = ns.GetDefaultFolder(constants.olFolderContacts)
        except AttributeError:
            contacts_folder = ns.GetDefaultFolder(10)
        print("Buzón {} no encontrado. Usando buzón por defecto.".format(MAILBOX_BUSCAR))

    import re

    def get_emails_from_contact(contact):
        emails = []
        for attr in ("Email1Address", "Email2Address", "Email3Address"):
            val = getattr(contact, attr, None)
            if val and isinstance(val, (str, unicode)) and val.strip():
                emails.append(val.strip())
        if not emails:
            for attr in ("Email1DisplayName", "Email2DisplayName", "Email3DisplayName"):
                val = getattr(contact, attr, None)
                if val and isinstance(val, (str, unicode)) and val.strip():
                    m = re.search(r'([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})', val)
                    if m:
                        emails.append(m.group(1))
        return emails

    def get_contacts_from_folder(folder):
        found = []
        try:
            items = folder.Items
            for i in range(1, items.Count + 1):
                try:
                    contact = items.Item(i)
                    if getattr(contact, "Class", 0) != 40:
                        continue
                    nombre = getattr(contact, "FullName", None) or ""
                    if not nombre or not nombre.strip():
                        nombre = (getattr(contact, "FirstName", "") or "") + " " + (getattr(contact, "LastName", "") or "")
                        nombre = nombre.strip() or "(Sin nombre)"
                    for email in get_emails_from_contact(contact):
                        if DOMINIO_FILTRO.lower() in email.lower():
                            found.append({"nombre": nombre.strip() or email, "email": email})
                            break
                except Exception:
                    continue
            for j in range(1, folder.Folders.Count + 1):
                sub = folder.Folders.Item(j)
                found.extend(get_contacts_from_folder(sub))
        except Exception:
            pass
        return found

    resultados = get_contacts_from_folder(contacts_folder)
    return resultados


def cargar_personas_existentes():
    """Carga el archivo personas.json actual"""
    if not os.path.exists(PERSONAS_FILE):
        return []
    try:
        with open(PERSONAS_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return []


def guardar_personas(personas):
    """Guarda la lista de personas en personas.json"""
    try:
        os.makedirs(ISSUES_DIR, exist_ok=True)
        with open(PERSONAS_FILE, "w", encoding="utf-8") as f:
            json.dump(personas, f, ensure_ascii=False, indent=2)
    except Exception as e:
        print("ERROR al guardar: " + str(e))
        return False
    return True


def main():
    print("Leyendo contactos de Outlook...")
    contactos = obtener_emails_arainco_desde_outlook()
    print("Encontrados {} contactos con @arainco.cl".format(len(contactos)))
    if contactos:
        for c in contactos[:5]:
            print("  Ejemplo: {} <{}>".format(c["nombre"], c["email"]))

    if not contactos:
        print("No hay contactos con direccion @arainco.cl en Outlook.")
        return

    existentes = cargar_personas_existentes()
    emails_existentes = {p.get("email", "").lower() for p in existentes}
    nombres_existentes = {p.get("nombre", "").lower() for p in existentes}

    agregados = 0
    for c in contactos:
        email_lower = (c.get("email") or "").lower()
        nombre_lower = (c.get("nombre") or "").lower()
        if email_lower not in emails_existentes and nombre_lower not in nombres_existentes:
            existentes.append(c)
            emails_existentes.add(email_lower)
            nombres_existentes.add(nombre_lower)
            agregados += 1
            print("  + Agregado: {} <{}>".format(c["nombre"], c["email"]))

    if agregados > 0:
        if guardar_personas(existentes):
            print("\nListo. Se agregaron {} personas al directorio.".format(agregados))
        else:
            print("\nError al guardar.")
    else:
        print("\nTodos los contactos ya estaban en el directorio.")


if __name__ == "__main__":
    main()
