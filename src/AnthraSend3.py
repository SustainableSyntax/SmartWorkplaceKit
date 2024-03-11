"""
Author: Hendrik Siemens
Date: 08/03/2024
Version: 3.0

Description:
    This automated script reads user data from an Excel file and sends
    personalized e-mails to the listed individuals in their respective
    languages.

    The e-mail content, including the subject and body, notifies recipients
    about the postponement of a planned change from the URL
    'job.takko.com' to 'application.takko.com'.

    Automated localization supports four distinct languages:
        - German (DE),
        - English (EN),
        - French (FR),
        - Dutch (NL).
    Special consideration is taken for French store managers with a tailored
    message.

    The script utilizes the 'win32com.client' package to interface with
    Microsoft Outlook and dispatch e-mails using the account configured
    in the default Outlook profile.

    It leverages 'pandas' to parse the Excel file and extract necessary
    user information to produce accurate mail content. Each recipient's
    e-mail is individually crafted according to their language preference
    indicated within the Excel data.

    To aid with clarity and maintenance, classes such as
    'User', 'Email', 'EmailContentGenerator', and 'MailSender' are utilized,
    encapsulating user details, e-mail construction, content generation logic,
    and e-mail dispatching functionality respectively. Error handling and user
    interaction features are embedded to enhance reliability an
    user-friendliness during the e-mail dispatch process.

Important:
    This script is tailored for execution on a Windows operating system with
    Microsoft Outlook installed and correctly set up. Execution within
    other environments or with different e-mail clients is not supported
    by the current implementation.

Usage:
    Ensure that the Outlook application is installed and the user has an
    active e-mail profile.

    Execute the script on a Windows machine where the Excel file with
    user data is accessible.

    The script will prompt the user for confirmation before dispatching
    e-mails.

    Progress feedback is provided during the sending process, and any
    errors encountered will be reported to the user.

Functions:
    main() - The primary function invoked when the script is run; responsible
                for the overall logic flow.
    generate_emails() - Reads user data from the Excel file, generates e-mail
                content, and constructs e-mail objects.
    _get_localized_content() - Produces subject and body text for an e-mail in
                the intended recipient's language.

Classes:
    User - A data structure holding details about a user such as:
                - username,
                - first name,
                - language,
                - salutation.
    Email - Represents an e-mail with fields for the recipient's:
                - address,
                - subject,
                - body.
    EmailContentGenerator - Encapsulates the logic for creating
                personalized, localized content for the e-mails.
    MailSender - Handles the sending of e-mail objects via the
                Outlook application.
    ErrorHandler - Manages error reporting by displaying messages
                to the user when an exception occurs.
    UIHandler - Manages user interaction elements, including
                confirmation prompts and progress bars.
"""

import win32com.client
import pandas as pd
import sys
from PyQt5.QtWidgets import QApplication, QMessageBox, QProgressBar


class Email:
    def __init__(self, to_email, subject, body):
        self.to_email = to_email
        self.subject = subject
        self.body = body


class User:
    def __init__(self, username, firstname, language, salutation):
        self.username = username
        self.firstname = firstname
        self.language = language
        self.salutation = salutation


class EmailContentGenerator:
    SALUTATIONS = {
        'MR': {'DE': "Lieber", 'EN': "Dear", 'FR': "Cher", 'NL': "Beste"},
        'MS': {'DE': "Liebe", 'EN': "Dear", 'FR': "Chère", 'NL': "Beste"},
        'MX': {'DE': "Liebe(r)", 'EN': "Dear", 'FR': "Cher(e)", 'NL': "Beste"},
    }

    def create_email_content(self, user: User):
        anrede_text = self.SALUTATIONS.get(
            user.salutation, self.SALUTATIONS['MX'])
        subject, body = self._get_localized_content(
            user.language, anrede_text[user.language], user.firstname)
        return Email(user.username, subject, body)

    def _get_localized_content(self, language, salutation_text, vorname):
        # Gemeinsame URL für alle Sprachen in einer Variablen halten
        current_url = ">>> job.takko.com <<<"
        planned_change_date = "08.03.2024"
        common_info = (
            f"die für den {planned_change_date} angesetzt war, vorerst aufgeschoben wird. Es wird keine Änderung geben, "
            "bis weitere Informationen bereitgestellt werden.\n\n"
            f"Bitte verwendet weiterhin die aktuelle URL {current_url} für den Zugriff auf d.vinci."
        )

        body_texts = {
            'DE': (
                f"{salutation_text} {vorname},\n\n"
                f"ich möchte euch darüber informieren, dass {common_info} Ich entschuldige mich "
                "für die Unannehmlichkeiten und danke euch für euer Verständnis und eure Flexibilität.\n\n"
                "Für Rückfragen stehe ich euch gerne zur Verfügung.\n\n"
                "Mit freundlichen Grüßen,\nHendrik Siemens"
            ),
            'EN': (
                f"{salutation_text} {vorname},\n\n"
                "I wish to inform you that the planned change of the URL from [job.takko.com] to [application.takko.com], which was "
                f"scheduled for {planned_change_date}, has been postponed until further notice. "
                f"Please continue to use the current URL {current_url} to access d.vinci. I apologize for any inconvenience and "
                "thank you for your understanding and flexibility.\n\n"
                "Should you have any questions, please do not hesitate to contact me.\n\n"
                "Kind regards,\nHendrik Siemens"
            ),
            'FR': (
                f"{salutation_text} {vorname},\n\n"
                "je tiens à vous informer que la modification prévue de l'URL de [job.takko.com] à [application.takko.com], prévue pour "
                f"le {planned_change_date}, est reportée jusqu'à nouvel ordre. "
                f"Veuillez donc continuer à utiliser l'URL actuelle {current_url} pour accéder à d.vinci. Je m'excuse pour les désagréments causés et "
                "vous remercie de votre compréhension et de votre flexibilité.\n\n"
                "En cas de questions, n'hésitez pas à me contacter.\n\n"
                "Cordialement,\nHendrik Siemens"
            ),
            'NL': (
                f"{salutation_text} {vorname},\n\n"
                "Ik wil jullie informeren dat de geplande wijziging van de URL van [job.takko.com] naar [application.takko.com], die gepland stond voor "
                f"{planned_change_date}, voorlopig is uitgesteld. Er worden geen wijzigingen aangebracht tot nadere informatie beschikbaar is.\n\n"
                f"Gelieve de huidige URL {current_url} te blijven gebruiken om toegang te krijgen tot d.vinci. Mijn excuses voor eventuele overlast en "
                "dank voor uw begrip en flexibilitet.\n\n"
                "Als u vragen heeft, neem dan gerust contact met mij op.\n\n"
                "Met vriendelijke groet,\nHendrik Siemens"
            )
        }

        subject_texts = {
            'DE': "Dringend: Aufschub der geplanten Änderung der URL für d.vinci",
            'EN': "Urgent: Postponement of the Planned URL Change for d.vinci",
            'FR': "Urgent: Report de la modification prévue de l'URL pour d.vinci",
            'NL': "Dringend: Uitstel van de geplande URL-wijziging voor d.vinci"
        }

        if language not in subject_texts or language not in body_texts:
            raise ValueError("Unsupported language")

        if language == 'FR' and "Responsable de magasin" in vorname:
            body_texts['FR'] = (
                "Bonjour à tous,\n\n"
                "Je tiens à vous informer que la modification prévue de l'URL de [job.takko.com] à [application.takko.com], "
                f"prévue pour le {planned_change_date}, est reportée jusqu'à nouvel ordre. {common_info} "
                "Je m'excuse pour les désagréments causés et vous remercie de votre compréhension et de votre flexibilité.\n\n"
                "En cas de questions, n'hésitez pas à me contacter.\n\n"
                "Cordialement,\nHendrik Siemens"
            )

        subject = subject_texts[language]
        body = body_texts[language]

        return subject, body


class MailSender:
    def __init__(self):
        self.outlook = win32com.client.Dispatch("Outlook.Application")

    def send_email(self, email):
        try:
            mail = self.outlook.CreateItem(0)
            mail.To = email.to_email
            mail.Subject = email.subject
            mail.Body = email.body
            mail.Send()
            print(f"E-Mail an {email.to_email} gesendet.")
        except Exception as e:
            raise e


class ErrorHandler:
    @staticmethod
    def display_error_message(to_email, e):
        msgBox = QMessageBox()
        msgBox.setIcon(QMessageBox.Warning)
        msgBox.setText(f"Fehler beim Versand an {to_email}: {e}")
        msgBox.setWindowTitle("Fehler")
        msgBox.exec_()


class UIHandler:
    @staticmethod
    def prompt_user_confirmation(count):
        app = QApplication(sys.argv)
        msgBox = QMessageBox()
        msgBox.setIcon(QMessageBox.Warning)
        msgBox.setText(f"{count} E-Mails werden versendet. Sind Sie sicher?")
        msgBox.setWindowTitle("Warnung")
        msgBox.setStandardButtons(QMessageBox.Ok | QMessageBox.Cancel)
        retval = msgBox.exec_()
        if retval == QMessageBox.Cancel:
            sys.exit("Versand abgebrochen.")
        app.exit()

    @staticmethod
    def show_progress(emails, success_count, error_count):
        progress_bar = QProgressBar()
        progress_bar.setMaximum(len(emails))
        progress_bar.show()

        for i, email in enumerate(emails):
            progress_bar.setValue(i + 1)
            QApplication.processEvents()

        # waits for the user to close the progress bar
        progress_bar.close()


def generate_emails(df):
    content_generator = EmailContentGenerator()
    emails = []
    for index, row in df.iterrows():
        user = User(
            row['Benutzername'],
            row['Vorname'],
            row['Sprache'].upper(),
            row['Anrede']
        )
        email = content_generator.create_email_content(user)
        emails.append(email)
    return emails


def main():
    try:
        df = pd.read_excel('Mappe2.xlsx')
        emails = generate_emails(df)

        mail_sender = MailSender()
        success_count = 0
        error_count = []

        UIHandler.prompt_user_confirmation(len(emails))

        for email in emails:
            try:
                mail_sender.send_email(email)
                success_count += 1
            except Exception as e:
                ErrorHandler.display_error_message(email.to_email, e)
                error_count.append(email.to_email)

        UIHandler.show_progress(emails, success_count, error_count)
        print(
            f"{success_count} Mails sent successfully.\n{len(error_count)} failed to send."
        )
        return 0
    except Exception as e:
        print(f"Error: {e}")
        return 1


if __name__ == "__main__":
    val = main()
    sys.exit(val)
