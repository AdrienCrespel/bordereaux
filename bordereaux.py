import os
import logging
from dotenv import load_dotenv
import imaplib
import email
from pathlib import Path
from datetime import datetime
import fitz
from PIL import Image
import img2pdf

# Configurer le logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# Charger les variables d'environnement
load_dotenv()

email_address = os.getenv("EMAIL_ADDRESS")
password = os.getenv("EMAIL_PASSWORD")
imap_server = os.getenv("IMAP_SERVER")

# Dictionnaire des transporteurs et leurs noms dans les emails
transporters = {
    "chronopost": "Chronopost",
    "vinted-go": "Vinted Go",
    "mondial-relay": "Mondial Relay",
    "relais-colis": "Relais Colis",
    "ups": "UPS Access",
    "colissimo": "La Poste"
}

# Règles de rognage pour chaque transporteur
crop_rules = {
    "chronopost": (425, 0, 840, 600),
    "vinted-go": (5, 5, 295, 415),
    "mondial-relay": (0, 0, 600, 400),
    "relais-colis": (0, 70, 410, 600),
    "ups": (5, 5, 330, 450),
    "colissimo": (0, 0, 415, 600),
    "autres": (0, 0, 600, 800)  # Règle par défaut
}

def connect_to_mailbox(email_address, password, imap_server="imap.gmail.com"):
    """Connexion à la boîte mail."""
    try:
        mail = imaplib.IMAP4_SSL(imap_server)
        mail.login(email_address, password)
        logging.info("Connexion à la boîte mail réussie.")
        return mail
    except Exception as e:
        logging.error(f"Erreur de connexion à la boîte mail: {e}")
        raise

def get_unread_emails_with_subject(mail, subject_keyword):
    """Récupérer les emails non lus avec un sujet donné."""
    try:
        mail.select("inbox")
        search_criteria = f'(UNSEEN SUBJECT "{subject_keyword}")'
        status, messages = mail.search(None, search_criteria)
        email_ids = messages[0].split()
        logging.info(f"Nombre de mails trouvés : {len(email_ids)}")
        return email_ids
    except Exception as e:
        logging.error(f"Erreur lors de la recherche des emails: {e}")
        return []

def get_transporter_from_email_body(email_body):
    """Déterminer le transporteur à partir du corps de l'email."""
    email_body_lower = email_body.lower()
    for transporter, keyword in transporters.items():
        if keyword.lower() in email_body_lower:
            return transporter
    return "autres"

def download_attachments_by_transporter(mail, unread_emails, date_folder):
    """Télécharger les pièces jointes en fonction du transporteur."""
    downloaded_pdfs = []

    for email_id in unread_emails:
        try:
            status, msg_data = mail.fetch(email_id, '(RFC822)')
            for response_part in msg_data:
                if isinstance(response_part, tuple):
                    msg = email.message_from_bytes(response_part[1])
                    email_body = ""

                    if msg.is_multipart():
                        for part in msg.walk():
                            content_type = part.get_content_type()
                            if content_type == "text/plain" or content_type == "text/html":
                                email_body = part.get_payload(decode=True).decode(part.get_content_charset(), errors="ignore")
                                break
                    else:
                        email_body = msg.get_payload(decode=True).decode(msg.get_content_charset(), errors="ignore")

                    transporter = get_transporter_from_email_body(email_body)
                    transporter_folder = date_folder / transporter
                    transporter_folder.mkdir(parents=True, exist_ok=True)

                    for part in msg.walk():
                        if part.get_content_maintype() == 'multipart':
                            continue
                        if part.get('Content-Disposition') is None:
                            continue
                        file_name = part.get_filename()
                        if file_name and file_name.lower().endswith('.pdf'):
                            file_path = transporter_folder / file_name
                            with file_path.open('wb') as f:
                                f.write(part.get_payload(decode=True))
                            logging.info(f"Pièce jointe téléchargée : {file_path}")
                            downloaded_pdfs.append(file_path)
        except Exception as e:
            logging.error(f"Erreur lors du téléchargement des pièces jointes pour l'email ID {email_id}: {e}")

    return downloaded_pdfs

def crop_shipping_document(pdf_path, transporter):
    """Rogner le document PDF en fonction du transporteur."""
    try:
        document = fitz.open(pdf_path)
        cropped_images_paths = []

        for page_num in range(len(document)):
            page = document.load_page(page_num)
            pix = page.get_pixmap()
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)

            # Appliquer le rognage
            cropped_img = img.crop(crop_rules.get(transporter, crop_rules["autres"]))

            # Appliquer une rotation de 90 degrés pour Mondial Relay
            if transporter == "mondial-relay":
                cropped_img = cropped_img.rotate(90, expand=True)

            # Construire le nom de fichier de sortie
            cropped_file_name = f"cropped_{pdf_path.stem}.png"
            cropped_file_path = pdf_path.parent / cropped_file_name

            cropped_img.save(cropped_file_path)
            logging.info(f"Image rognée enregistrée : {cropped_file_path}")
            cropped_images_paths.append(cropped_file_path)

        return cropped_images_paths
    except Exception as e:
        logging.error(f"Erreur lors du rognage du PDF {pdf_path}: {e}")
        return []

def merge_images_to_pdf(image_paths, output_pdf_path):
    """Fusionner les images en un seul PDF sans agrandir les images."""
    try:
        # Convertir les images en PDF sans les agrandir
        with open(output_pdf_path, "wb") as f:
            f.write(img2pdf.convert(image_paths))

        logging.info(f"PDF fusionné enregistré : {output_pdf_path}")
    except Exception as e:
        logging.error(f"Erreur lors de la fusion des images en PDF: {e}")

def main():
    mail = connect_to_mailbox(email_address, password)
    subject_keyword = "Bordereau d'envoi Vinted"
    unread_emails = get_unread_emails_with_subject(mail, subject_keyword)

    now = datetime.now()
    date_folder_name = now.strftime("%Y-%m-%d_%H-%M-%S")
    base_path = Path("bordereaux")
    date_folder = base_path / date_folder_name

    downloaded_pdfs = download_attachments_by_transporter(mail, unread_emails, date_folder)

    cropped_pdfs = []
    for pdf_path in downloaded_pdfs:
        transporter = pdf_path.parent.name
        cropped_images = crop_shipping_document(pdf_path, transporter)
        cropped_pdfs.append(cropped_images)

    all_cropped_images = [img for cropped_set in cropped_pdfs for img in cropped_set]
    output_pdf_path = date_folder / "bordereaux.pdf"
    merge_images_to_pdf(all_cropped_images, output_pdf_path)

if __name__ == "__main__":
    main()
