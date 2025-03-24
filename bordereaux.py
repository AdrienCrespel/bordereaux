import os
import re
import logging
from dotenv import load_dotenv
import imaplib
import email
from email.header import decode_header
from pathlib import Path
from datetime import datetime
from reportlab.pdfgen import canvas
from PyPDF2 import PdfReader, PdfWriter, PdfMerger
import io

# Configurer le logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# Charger les variables d'environnement
load_dotenv()

email_address = os.getenv("EMAIL_ADDRESS")
password = os.getenv("EMAIL_PASSWORD")
imap_server = os.getenv("IMAP_SERVER")

# Structure centralisée pour les transporteurs
transporters_info = {
    "chronopost": {
        "name": "Chronopost",
        "id_position": (510, 510),
        "cut_rules": lambda width, height: ((width / 2, 0), (width, height), None)
    },
    "vinted-go": {
        "name": "Vinted Go",
        "id_position": (20, 760),
        "cut_rules": lambda width, height: ((0, height / 2), (width / 2, height), None)
    },
    "mondial-relay": {
        "name": "Mondial Relay",
        "id_position": (45, 750),
        "cut_rules": lambda width, height: ((0, height / 2), (width, height), "rotate")
    },
    "relais-colis": {
        "name": "Relais Colis",
        "id_position": (100, 500),
        "cut_rules": lambda width, height: ((0, 0), (width / 2, height), None)
    },
    "ups": {
        "name": "UPS Access",
        "id_position": (100, 560),
        "cut_rules": lambda width, height: ((0, height * 0.233), (width * 0.401, height), None)
    },
    "colissimo": {
        "name": "La Poste",
        "id_position": (50, 510),
        "cut_rules": lambda width, height: ((0, 0), (width / 2, height), None)
    },
    "autres": {
        "name": "Autres",
        "id_position": (50, 500),
        "cut_rules": lambda width, height: ((0, 0), (width, height), None)
    }
}

def get_transporter_name(transporter):
    """Obtenir le nom du transporteur."""
    return transporters_info.get(transporter, {}).get("name", "Inconnu")

def get_id_position(transporter, width, height):
    """Obtenir les coordonnées pour l'ID du transporteur."""
    return transporters_info.get(transporter, {}).get("id_position", (0, 0))

def get_cut_coordinates(transporter, width, height):
    """Obtenir les coordonnées de découpe pour le transporteur."""
    cut_rules = transporters_info.get(transporter, {}).get("cut_rules", lambda w, h: ((0, 0), (w, h), None))
    return cut_rules(width, height)

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

def extract_id_from_subject(subject):
    """Extraire l'ID du sujet de l'email."""
    # Expression régulière pour extraire l'ID (séquence alphanumérique)
    match = re.search(r'pour ([\w-]+) -', subject)
    if match:
        return match.group(1)
    return "UNKNOWN"

def get_unread_emails_with_subject(mail, subject_keyword):
    """Récupérer les emails non lus avec un sujet donné et extraire les IDs."""
    try:
        mail.select("inbox")
        search_criteria = f'(UNSEEN SUBJECT "{subject_keyword}")'
        status, messages = mail.search(None, search_criteria)
        email_ids = messages[0].split()

        email_subjects_and_ids = []
        for email_id in email_ids:
            status, msg_data = mail.fetch(email_id, '(RFC822)')
            for response_part in msg_data:
                if isinstance(response_part, tuple):
                    msg = email.message_from_bytes(response_part[1])
                    decoded_subject, encoding = decode_header(msg["Subject"])[0]
                    if isinstance(decoded_subject, bytes):
                        # Si le sujet est en bytes, le décoder
                        decoded_subject = decoded_subject.decode(encoding if encoding else "utf-8")
                    email_subject = decoded_subject
                    email_id_extracted = extract_id_from_subject(email_subject)
                    email_subjects_and_ids.append((email_id, email_id_extracted))

        return email_subjects_and_ids
    except Exception as e:
        logging.error(f"Erreur lors de la recherche des emails: {e}")
        return []

def get_transporter_from_email_body(email_body):
    """Déterminer le transporteur à partir du corps de l'email."""
    email_body_lower = email_body.lower()
    for transporter, info in transporters_info.items():
        if info["name"].lower() in email_body_lower:
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

def add_id_to_pdf(input_pdf_path, output_pdf_path, id_text, transporter):
    """Ajouter l'ID sur le PDF en fonction du transporteur."""
    try:
        # Lire le PDF existant
        reader = PdfReader(input_pdf_path)
        writer = PdfWriter()

        # Parcourir chaque page du PDF (ici, une seule page)
        page = reader.pages[0]
        packet = io.BytesIO()
        width = page.mediabox.width
        height = page.mediabox.height

        print(width, height)

        # Obtenir les coordonnées pour le transporteur
        x, y = get_id_position(transporter, width, height)
        
        print(x, y)

        # Créer un canevas pour dessiner sur la page
        can = canvas.Canvas(packet)
        can.drawString(x, y, f"SKU: {id_text}")  # Positionner l'ID selon le transporteur
        can.save()

        # Mouvement vers le début du flux
        packet.seek(0)
        new_pdf = PdfReader(packet)
        new_page = new_pdf.pages[0]

        # Ajouter le contenu du canevas à la page existante
        page.merge_page(new_page)
        writer.add_page(page)

        # Enregistrer le PDF modifié
        with open(output_pdf_path, "wb") as f:
            writer.write(f)

        logging.info(f"ID ajouté au PDF : {output_pdf_path}")
    except Exception as e:
        logging.error(f"Erreur lors de l'ajout de l'ID au PDF: {e}")

def crop_shipping_document(pdf_path, transporter, extracted_id):
    """Rogner le document PDF et ajouter l'ID en fonction du transporteur."""
    try:
        with open(pdf_path, 'rb') as file:
            reader = PdfReader(file)
            writer = PdfWriter()            
            page = reader.pages[0]
            width = page.mediabox.width
            height = page.mediabox.height

            # Obtenir les coordonnées de découpe en fonction du transporteur
            lower_left, upper_right, rotation = get_cut_coordinates(transporter, width, height)

             # Appliquer les coordonnées de découpe à la page
            page.mediabox.lower_left = (int(lower_left[0]), int(lower_left[1]))
            page.mediabox.upper_right = (int(upper_right[0]), int(upper_right[1]))

            # Appliquer la rotation si nécessaire
            if rotation == "rotate":
                page.rotate(90)

            # Ajouter la première page découpée au fichier de sortie
            writer.add_page(page)

            # Sauvegarder le nouveau PDF
            output_pdf = pdf_path.parent / f"cropped_{pdf_path.name}"
            with open(output_pdf, 'wb') as output_file:
                writer.write(output_file)
                logging.info(f"Le pdf rogné a correctement été généré : {output_file.name}")
            
            # Ajouter l'ID sur le PDF
            add_id_to_pdf(output_pdf, output_pdf, extracted_id, transporter)
            
            return output_pdf
    except Exception as e:
        logging.error(f"Erreur lors du rognage du PDF {pdf_path}: {e}")
        return None

def merge_pdfs(pdf_paths, output_pdf_path):
    """Fusionner les PDF."""
    try:
        merger = PdfMerger()
        for pdf in pdf_paths:
            merger.append(pdf)

        # Écrire le fichier fusionné
        merger.write(output_pdf_path)
        merger.close()
        logging.info(f"PDF fusionné enregistré : {output_pdf_path}")
    except Exception as e:
        logging.error(f"Erreur lors de la fusion des PDF: {e}")

def save_ids_to_file(ids, output_file_path):
    """Enregistrer les IDs extraits dans un fichier texte."""
    try:
        with open(output_file_path, 'w') as file:
            for id_value in ids:
                file.write(f"{id_value}\n")
        logging.info(f"IDs enregistrés dans le fichier : {output_file_path}")
    except Exception as e:
        logging.error(f"Erreur lors de l'enregistrement des IDs dans le fichier : {e}")

def main():
    mail = connect_to_mailbox(email_address, password)
    subject_keyword = "Bordereau d'envoi Vinted"
    unread_emails_with_ids = get_unread_emails_with_subject(mail, subject_keyword)

    now = datetime.now()
    date_folder_name = now.strftime("%Y-%m-%d_%H-%M-%S")
    base_path = Path("bordereaux")
    date_folder = base_path / date_folder_name

    # Télécharger les PDFs et associer chaque PDF à son ID
    downloaded_pdfs_with_ids = []
    for email_id, extracted_id in unread_emails_with_ids:
        downloaded_pdfs = download_attachments_by_transporter(mail, [email_id], date_folder)
        for pdf_path in downloaded_pdfs:
            downloaded_pdfs_with_ids.append((pdf_path, extracted_id))

    cropped_pdfs_with_ids = []
    for pdf_path, extracted_id in downloaded_pdfs_with_ids:
        transporter = pdf_path.parent.name
        cropped_pdf = crop_shipping_document(pdf_path, transporter, extracted_id)
        if cropped_pdf:
            cropped_pdfs_with_ids.append((cropped_pdf, extracted_id))

    # Vérifier s'il y a des PDFs à fusionner
    if cropped_pdfs_with_ids:
        output_pdf_path = date_folder / "bordereaux.pdf"
        merge_pdfs([pdf for pdf, _ in cropped_pdfs_with_ids], output_pdf_path)
    else:
        logging.info("Aucun PDF à fusionner.")

    # Extraire les IDs et les enregistrer dans un fichier texte
    extracted_ids = [extracted_id for _, extracted_id in unread_emails_with_ids if extracted_id is not None]

    # Vérifier s'il y a des IDs à enregistrer
    if extracted_ids:
        ids_file_path = date_folder / "bordereaux.txt"
        save_ids_to_file(extracted_ids, ids_file_path)
    else:
        logging.info("Aucun ID à enregistrer.")
    

if __name__ == "__main__":
    main()
