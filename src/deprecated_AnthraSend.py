import time
import win32com.client
import pandas as pd

outlook = win32com.client.Dispatch("Outlook.Application")

excel_file = 'Mappe2.xlsx'
df = pd.read_excel(excel_file)

for index, row in df.iterrows():
    to_email = row['Benutzername']
    anrede = row['Anrede']
    nachname = row['Nachname']
    vorname = row['Vorname']
    sprache = row['Sprache'].upper()

    if anrede == 'MR':
        anrede_text_de = "Lieber"
        anrede_text_en = "Dear"
        anrede_text_fr = "Cher"
        anrede_text_nl = "Beste"
    elif anrede == 'MS':
        anrede_text_de = "Sehr geehrte Frau"
        anrede_text_en = "Dear Madam"
        anrede_text_fr = "Madame"
        anrede_text_nl = "Geachte mevrouw"
    elif anrede == 'MX':
        anrede_text_de = "Liebe"
        anrede_text_en = "Dear"
        anrede_text_fr = ""
        anrede_text_nl = "Beste"
    else:
        anrede_text_de = "Liebe(r)"
        anrede_text_en = "Dear"
        anrede_text_fr = ""
        anrede_text_nl = "Beste"

    subject = ""
    body = ""

    if sprache == 'DE':
        subject = "Wichtige Information: Umstellung der URL von job.takko.com auf application.takko.com am 08.03.2024"
        body = (
            f"Hallo {vorname},\n\n"
            "wir möchten euch über eine bevorstehende Änderung informieren, die ab dem 08.03.2024 wirksam wird.\n\nDie URL von [job.takko.com] wird zu [application.takko.com] geändert.\n\nBitte stellt sicher, dass ihr eure Lesezeichen oder Favoriten aktualisiert und ab diesem Datum die neue URL verwendet, um auf d.vinci zuzugreifen.\nSolltet ihr Probleme bei der Anmeldung haben, zögert nicht, uns zu kontaktieren. Wir stehen euch gerne zur Verfügung, um bei eventuellen Anpassungen behilflich zu sein.\n\n"
            "Mit freundlichen Grüßen,\nHendrik Siemens"
        )
    elif sprache == 'EN':
        subject = "Important Information: Change of URL from job.takko.com to application.takko.com on 08.03.2024"
        body = (
            f"Hello {vorname},\n\n"
            "We would like to inform you about an upcoming change that will take effect on 08.03.2024.\n\nThe URL from [job.takko.com] will be changed to [application.takko.com].\n\nPlease ensure that you update your bookmarks or favorites and use the new URL from this date onwards to access d.vinci.\nIf you encounter any issues with logging in, feel free to contact us. We are here to assist you with any necessary adjustments.\n\n"
            "Best regards,\nHendrik Siemens"
        )
    elif sprache == 'FR':
        if "Responsable de magasin" in vorname:
            subject = "Information importante : Changement de l'URL de job.takko.com en application.takko.com le 08.03.2024"
            body = (
                f"Bonjour à tous,\n\n"
                "Nous tenons à vous informer d'un changement à venir qui entrera en vigueur le 08.03.2024.\n\nL'URL de job.takko.com sera modifiée en application.takko.com.\n\nVeuillez vous assurer de mettre à jour vos favoris et d'utiliser la nouvelle URL à partir de cette date pour accéder à d.vinci.\nSi vous rencontrez des problèmes de connexion, n'hésitez pas à nous contacter. Nous sommes là pour vous aider à apporter les ajustements nécessaires.\n\n"
                "Cordialement,\nHendrik Siemens"
            )
        else:
            subject = "Information importante : Changement de l'URL de job.takko.com en application.takko.com le 08.03.2024"
            body = (
                f"Bonjour {vorname},\n\n"
                "Nous tenons à vous informer d'un changement à venir qui entrera en vigueur le 08.03.2024.\n\nL'URL de job.takko.com sera modifiée en application.takko.com.\n\nVeuillez vous assurer de mettre à jour vos favoris et d'utiliser la nouvelle URL à partir de cette date pour accéder à d.vinci.\nSi vous rencontrez des problèmes de connexion, n'hésitez pas à nous contacter. Nous sommes là pour vous aider à apporter les ajustements nécessaires.\n\n"
                "Cordialement,\nHendrik Siemens"
            )

    elif sprache == 'NL':
        subject = "Belangrijke informatie: Wijziging van de URL van job.takko.com naar application.takko.com op 08.03.2024"
        body = (
            f"Beste {vorname},\n\n"
            "Wij willen jullie graag informeren over een aankomende verandering die vanaf 08.03.2024 van kracht zal zijn.\n\nDe URL van [job.takko.com] zal worden gewijzigd naar [application.takko.com].\n\nZorg ervoor dat je je bladwijzers of favorieten bijwerkt en vanaf deze datum de nieuwe URL gebruikt om toegang te krijgen tot d.vinci.\nMocht je problemen ondervinden bij het inloggen, neem dan gerust contact met ons op. Wij staan voor jullie klaar om te helpen bij eventuele aanpassingen.\n\n"
            "Met vriendelijke groet,\nHendrik Siemens"
        )

    mail = outlook.CreateItem(0)
    mail.To = to_email
    mail.Subject = subject
    mail.Body = body
    
    mail.Send()
    
    print(f"E-Mail an {to_email} gesendet")
    time.sleep(5)
