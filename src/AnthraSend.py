"""
Author: Hendrik Siemens
Date: 04/03/2024
Version: 2.0

Description:
    This script reads an Excel file and sends e-mails to the respective
    people listed. The content of the e-mail is generated based on the
    data from the Excel file and is customized to the language of the
    addressee.

    In this particular case the script was written to inform all users
    about an upcoming change of the URL from

    job.takko.com

    to

    application.takko.com.

    The mail body is written in four languages:
        - Germam (DE)
        - English (EN)
        - French (FR)
        - Dutch (NL)

    The script uses the win32com.client package to interact with Microsoft
    Outlook and send mails using the default Outlook profile and settings.
    It uses pandas to read the Excel file and generate the required mail
    content based on the data found in the Excel file.

    The script iterates through each e-mail address found in the Excel
    file and sends a mail to each e-mail address with the respective
    content in the correct language.

    Note:
        The script is written to run on a Windows machine with a local
        installation of Outlook.
"""

import sys
import time
import win32com.client
import pandas as pd

from PyQt5.QtWidgets import QApplication, QMessageBox, QProgressBar


def read_email_data(excel_file):
    return pd.read_excel(excel_file)


def generate_emails(df):
    emails = []

    for index, row in df.iterrows():
        to_email = row['Benutzername']
        anrede = row['Anrede']
        _ = row['Nachname']  # currently not used
        vorname = row['Vorname']
        sprache = row['Sprache'].upper()

        if anrede == 'MR':
            anrede_text_de = "Lieber"
            anrede_text_en = "Dear"
            anrede_text_fr = "Cher"
            anrede_text_nl = "Beste"
        elif anrede == 'MS':
            anrede_text_de = "Liebe"
            anrede_text_en = "Dear"
            anrede_text_fr = "Chère"
            anrede_text_nl = "Beste"
        elif anrede == 'MX':
            anrede_text_de = "Liebe(r)"
            anrede_text_en = "Dear"
            anrede_text_fr = "Cher(e)"
            anrede_text_nl = "Beste"
        else:
            anrede_text_de = "Liebe(r)"
            anrede_text_en = "Dear"
            anrede_text_fr = "Cher(e)"
            anrede_text_nl = "Beste"

        subject = ""
        body = ""

        if sprache == 'DE':
            subject = \
                "Dringend: Aufschub der geplanten Änderung der URL für d.vinci"
            body = (
                f"{anrede_text_de} {vorname},\n\n"
                "ich möchte euch darüber informieren, dass die geplante Änderung der URL von [job.takko.com] zu [application.takko.com], "
                "die für den 08.03.2024 angesetzt war, vorerst aufgeschoben wird. Es wird keine Änderung geben, "
                "bis weitere Informationen bereitgestellt werden.\n\n"
                "Bitte verwendet weiterhin die aktuelle URL >>> job.takko.com <<< für den Zugriff auf d.vinci. Ich entschuldige mich für die Unannehmlichkeiten und "
                "danke euch für euer Verständnis und eure Flexibilität.\n\n"
                "Für Rückfragen stehe ich euch gerne zur Verfügung.\n\n"
                "Mit freundlichen Grüßen,\nHendrik Siemens"
            )
        elif sprache == 'EN':
            subject = \
                "Urgent: Postponement of the Planned URL Change for d.vinci"
            body = (
                f"{anrede_text_en} {vorname},\n\n"
                "I wish to inform you that the planned change of the URL from [job.takko.com] to [application.takko.com], which was "
                "scheduled for 08.03.2024, has been postponed until further notice. There will be no changes made until more information is provided.\n\n"
                "Please continue to use the current URL >>> job.takko.com <<< to access d.vinci. I apologize for any inconvenience and "
                "thank you for your understanding and flexibility.\n\n"
                "Should you have any questions, please do not hesitate to contact me.\n\n"
                "Kind regards,\nHendrik Siemens"
            )
        elif sprache == 'FR':
            if "Responsable de magasin" in vorname:
                subject = "Urgent : Report de la modification prévue de l'URL pour d.vinci"
                body = (
                    "Bonjour à tous,\n\n"
                    "Je tiens à vous informer que la modification prévue de l'URL de [job.takko.com] à [application.takko.com], prévue pour le 08.03.2024, "
                    "est reportée jusqu'à nouvel ordre. Aucun changement ne sera effectué jusqu'à la communication de nouvelles informations.\n\n"
                    "Veuillez donc continuer à utiliser l'URL actuelle >>> job.takko.com <<< pour accéder à d.vinci. Je m'excuse pour les désagréments causés et "
                    "vous remercie de votre compréhension et de votre flexibilité.\n\n"
                    "En cas de questions, n'hésitez pas à me contacter.\n\n"
                    "Cordialement,\nHendrik Siemens"
                )
            else:
                subject = "Urgent : Report de la modification prévue de l'URL pour d.vinci"
                body = (
                    f"{anrede_text_fr} {vorname},\n\n"
                    "je tiens à vous informer que la modification prévue de l'URL de [job.takko.com] à [application.takko.com], prévue pour le 08.03.2024, est "
                    "reportée jusqu'à nouvel ordre. Aucun changement ne sera effectué jusqu'à la communication de nouvelles informations.\n\n"
                    "Veuillez donc continuer à utiliser l'URL actuelle >>> job.takko.com <<< pour accéder à d.vinci. Je m'excuse pour les désagréments causés et "
                    "vous remercie de votre compréhension et de votre flexibilité.\n\n"
                    "En cas de questions, n'hésitez pas à me contacter.\n\n"
                    "Cordialement,\nHendrik Siemens"
                )

        elif sprache == 'NL':
            subject = \
                "Dringend: Uitstel van de geplande URL-wijziging voor d.vinci"
            body = (
                f"{anrede_text_nl} {vorname},\n\n"
                "Ik wil jullie informeren dat de geplande wijziging van de URL van [job.takko.com] naar [application.takko.com], die gepland stond voor 08.03.2024, "
                "voorlopig is uitgesteld. Er worden geen wijzigingen aangebracht tot nadere informatie beschikbaar is.\n\n"
                "Gelieve de huidige URL >>> job.takko.com <<< te blijven gebruiken om toegang te krijgen tot d.vinci. Mijn excuses voor eventuele overlast en "
                "dank voor uw begrip en flexibilitet.\n\n"
                "Als u vragen heeft, neem dan gerust contact met mij op.\n\n"
                "Met vriendelijke groet,\nHendrik Siemens"
            )

        email = {
            'to_email': to_email,
            'subject': subject,
            'body': body
        }

        emails.append(email)


def print_email_addresses(emails):
    mail_adresses_count = len(emails)
    print(f"Anzahl der Mail-Adressen: {mail_adresses_count}")
    max_length = max(len(email['to_email']) for email in emails)
    for i, email in enumerate(emails):
        print(
            f"E-Mail-Adresse {i + 1}: {email['to_email']: <{max_length}}", end=", ")
        if (i + 1) % 4 == 0:
            print()


def prompt_user_confirmation(count):
    _ = QApplication(sys.argv)
    msgBox = QMessageBox()
    msgBox.setIcon(QMessageBox.Warning)
    msgBox.setText(f"{count} E-Mails werden versendet. Sind Sie sicher?")
    msgBox.setWindowTitle("Warnung")
    msgBox.setStandardButtons(QMessageBox.Ok | QMessageBox.Cancel)
    msgBox.buttonClicked.connect(lambda button: sys.exit(
        "Mails not sent") if button.text() == "Cancel" else None)
    msgBox.exec_()


def send_emails(emails, outlook):
    progress_bar = QProgressBar()
    progress_bar.setMaximum(len(emails))
    progress_bar.show()

    success_count = 0
    error_count = []
    mail_send_delay = 5

    for i, email in enumerate(emails):
        try:
            mail = outlook.CreateItem(0)
            mail.To = email['to_email']
            mail.Subject = email['subject']
            mail.Body = email['body']

            mail.Send()

            print(f"E-Mail an {email['to_email']} gesendet")
            time.sleep(mail_send_delay)

            progress_bar.setValue(i + 1)
            success_count += 1
        except Exception as e:
            # Don't interrupt script execution, just log and continue
            display_error_message(email['to_email'], e)
            print(error_message(email['to_email'], e))
            error_count.append(email['to_email'])
    return success_count, error_count


def error_message(to_email, e):
    return f"Fehler beim Versand an {to_email}: {e}"


def display_error_message(to_email, e):
    msgBox = QMessageBox()
    msgBox.setIcon(QMessageBox.Warning)
    msgBox.setText(error_message(to_email, e))
    msgBox.setWindowTitle("Fehler")
    msgBox.exec_()


def display_success_message(success_count, total_count, error_count):
    success_message = \
        f"{success_count} out of {total_count} emails sent successfully" \
        f"\nwith {len(error_count)} errors:\n" \
        f"{error_count}" if error_count else \
        f"All {total_count} emails were sent successfully!"

    msgBox = QMessageBox()
    msgBox.setIcon(QMessageBox.Information)
    msgBox.setText(success_message)
    msgBox.setWindowTitle("Success")
    msgBox.exec_()

    print(success_message)
    return success_message


def main():
    outlook = win32com.client.Dispatch("Outlook.Application")
    df = read_email_data('Mappe2.xlsx')
    emails = generate_emails(df)

    print_email_addresses(emails)
    prompt_user_confirmation(len(emails))

    success_count, error_count = send_emails(emails, outlook)
    sys.exit(display_success_message(success_count, len(emails), error_count))


# Damit das Skript ausgeführt wird, wenn es direkt aufgerufen wird.
if __name__ == "__main__":
    main()
