import win32com.client

# Definieren des Strings, nach dem im Body der E-Mails gesucht werden soll
search_string = "Unzustellbar:"


def list_undeliverable_emails():
    # Verbindung zu Outlook herstellen
    outlook = win32com.client.Dispatch(
        "Outlook.Application").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder(6)  # 6 steht für den Posteingang

    # Liste für nicht zustellbare Mailadressen
    undeliverable_emails = []

    # Alle Nachrichten im Posteingang durchsuchen
    messages = inbox.Items
    for message in messages:
        if search_string in message.Subject:
            print(message)

    return undeliverable_emails


# Ausführung der Funktion und Ausgabe der Liste
undelivered_list = list_undeliverable_emails()

# Ausgabe der Anzahl der Rückläufer
print(f"Anzahl der Rückläufer: {len(undelivered_list)}")

# Ausgabe der gesammelten E-Mail-Adressen
print("Liste der E-Mail-Adressen mit Zustellungsfehlern:")
for email in undelivered_list:
    print(email)

# Denken Sie daran, dass dieses Skript Anpassungen benötigen könnte, abhängig von der genauen Struktur der Rückläufer-Nachrichten in Ihrem Outlook Postfach.
