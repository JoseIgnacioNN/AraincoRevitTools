# -*- coding: utf-8 -*-
"""
Prueba de envio de correo.
Usa las credenciales de recordatorios_config.json y recordatorios_config.local.json
Envia un correo de prueba a jose.nunez@arainco.cl

Ejecutar: python enviar_prueba_email.py
"""

import os
import sys
import json
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.utils import formataddr

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH = os.path.join(SCRIPT_DIR, "recordatorios_config.json")
LOCAL_CONFIG_PATH = os.path.join(SCRIPT_DIR, "recordatorios_config.local.json")


def _deep_merge(base, override):
    result = dict(base)
    for k, v in override.items():
        if k in result and isinstance(result[k], dict) and isinstance(v, dict):
            result[k] = _deep_merge(result[k], v)
        else:
            result[k] = v
    return result


def _load_config():
    if not os.path.exists(CONFIG_PATH):
        print("ERROR: No se encontro recordatorios_config.json")
        sys.exit(1)
    with open(CONFIG_PATH, "r", encoding="utf-8") as f:
        cfg = json.load(f)
    if os.path.isfile(LOCAL_CONFIG_PATH):
        try:
            with open(LOCAL_CONFIG_PATH, "r", encoding="utf-8") as f:
                local_cfg = json.load(f)
            cfg = _deep_merge(cfg, local_cfg)
        except Exception:
            pass
    return cfg


def main():
    cfg = _load_config()
    smtp_cfg = cfg.get("smtp", {})
    rem_cfg = cfg.get("remitente", {})

    host = smtp_cfg.get("server", "smtp.gmail.com")
    port = int(smtp_cfg.get("port", 587))
    use_tls = smtp_cfg.get("use_tls", True)
    use_ssl = smtp_cfg.get("use_ssl", port == 465)
    user = smtp_cfg.get("user", "")
    password = smtp_cfg.get("password", "")
    if not password:
        password = os.environ.get("RECORDATORIOS_SMTP_PASSWORD", "")

    from_email = rem_cfg.get("email", "") or user
    from_name = rem_cfg.get("nombre", "Arainco Notificaciones")
    to_email = "jose.nunez@arainco.cl"

    if not user or not password:
        print("ERROR: Falta smtp.user o smtp.password en la config")
        sys.exit(1)

    subject = "[PRUEBA] Recordatorios - Test de envio"
    body_plain = "Este es un correo de prueba del sistema de recordatorios de incidencias BIM."
    body_html = "<p>Este es un correo de <strong>prueba</strong> del sistema de recordatorios de incidencias BIM.</p>"

    msg = MIMEMultipart("alternative")
    msg["Subject"] = subject
    msg["From"] = formataddr((from_name, from_email))
    msg["To"] = to_email
    msg.attach(MIMEText(body_plain, "plain", "utf-8"))
    msg.attach(MIMEText(body_html, "html", "utf-8"))

    print("Enviando correo de prueba...")
    print("  De: {} <{}>".format(from_name, from_email))
    print("  Para: {}".format(to_email))
    print("  Servidor: {}:{}".format(host, port))

    try:
        if use_ssl or port == 465:
            conn = smtplib.SMTP_SSL(host, port)
        else:
            conn = smtplib.SMTP(host, port)
            if use_tls:
                conn.starttls()
        conn.login(user, password)
        conn.sendmail(from_email, [to_email], msg.as_string())
        conn.quit()
        print("OK: Correo enviado correctamente.")
    except Exception as e:
        print("ERROR: {}".format(e))
        sys.exit(1)


if __name__ == "__main__":
    main()
