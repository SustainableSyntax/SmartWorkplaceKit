import win32com.client

# Outlook-Anwendung öffnen
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

# Posteingangsordner auswählen
inbox = outlook.GetDefaultFolder(6)  # 6 steht für den Posteingangsordner

# Alle E-Mails im Posteingang auflisten
messages = inbox.Items
for message in messages:
    if "Unzustellbar:" in message.Subject:
        print(f"Rückläufer-Mail mit dem Betreff: {message.Subject}")
        print(message.Body.encode('utf-8'))

        body = message.Body
        if "Diagnoseinformationen für Administratoren:" in body:
            start_index = body.index(
                "Diagnoseinformationen für Administratoren:")
            error_info = body[start_index:]
            print("Fehlerdiagnoseinformationen:")
            print(error_info)

# Outlook-Anwendung beenden
del outlook
