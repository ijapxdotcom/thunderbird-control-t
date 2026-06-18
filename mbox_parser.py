import mailbox
import os
import email
import email.utils
import hashlib
from datetime import datetime, timedelta, timezone
from email.header import decode_header
import pypdf

def decode_mime_words(s):
    if not s:
        return ""
    try:
        parts = decode_header(s)
        decoded = []
        for part, charset in parts:
            if isinstance(part, bytes):
                decoded.append(part.decode(charset or 'utf-8', errors='replace'))
            else:
                decoded.append(part)
        return "".join(decoded)
    except Exception:
        return s

def generate_task_id(message_id):
    if not message_id:
        return hashlib.sha256(os.urandom(16)).hexdigest()[:16]
    clean_id = message_id.strip("<> \t\n\r")
    return hashlib.sha256(clean_id.encode('utf-8')).hexdigest()[:16]

def get_message_date(msg):
    date_str = msg.get('date')
    if not date_str:
        return None
    try:
        dt = email.utils.parsedate_to_datetime(date_str)
        if dt.tzinfo is None:
            dt = dt.replace(tzinfo=timezone.utc)
        return dt.astimezone(timezone.utc)
    except Exception:
        return None

def extract_pdf_text(pdf_path):
    text = ""
    try:
        reader = pypdf.PdfReader(pdf_path)
        for page_num, page in enumerate(reader.pages):
            page_text = page.extract_text()
            if page_text:
                text += f"--- PDF Page {page_num + 1} ---\n{page_text}\n"
    except Exception as e:
        print(f"Error reading PDF {pdf_path}: {e}")
    return text

def parse_mbox(mbox_path, window_hours=36, secure_attachments_dir="secure_attachments"):
    """
    Parses a local Thunderbird MBOX file and extracts emails within the specified hour window.
    Saves PDF attachments to the secure directory and extracts their text content.
    """
    if not os.path.exists(mbox_path):
        print(f"Mailbox path does not exist: {mbox_path}")
        return []

    print(f"Opening mailbox: {mbox_path} (Size: {os.path.getsize(mbox_path) / (1024*1024):.2f} MB)")
    
    mbox = mailbox.mbox(mbox_path)
    num_messages = len(mbox)
    print(f"Total messages in mailbox: {num_messages}")

    now_utc = datetime.now(timezone.utc)
    cutoff_time = now_utc - timedelta(hours=window_hours)
    print(f"Ingesting emails received after: {cutoff_time.isoformat()}")

    matching_emails = []
    
    # Process from newest to oldest. Stop if we hit a consecutive block of old emails.
    # Thunderbird stores newer emails towards the end of the file.
    consecutive_old_emails = 0
    max_consecutive_old = 15  # safety margin for out-of-order imports
    
    keys = mbox.keys()
    for key in reversed(keys):
        msg = mbox[key]
        msg_date = get_message_date(msg)
        
        if not msg_date:
            continue
            
        if msg_date < cutoff_time:
            consecutive_old_emails += 1
            if consecutive_old_emails >= max_consecutive_old:
                # We reached older emails, we can safely stop scanning backward.
                print(f"Reached {max_consecutive_old} consecutive emails older than the window. Stopping scan.")
                break
            continue
            
        # Reset counter since we found a valid email (could be slightly out of chronological order)
        consecutive_old_emails = 0
        
        message_id = msg.get('message-id')
        task_id = generate_task_id(message_id)
        
        subject = decode_mime_words(msg.get('subject', ''))
        sender = decode_mime_words(msg.get('from', ''))
        
        # Extract pure email address from From header (e.g. "Name <user@domain>" → "user@domain")
        _, sender_email = email.utils.parseaddr(sender)
        sender_email = sender_email.strip() if sender_email else ""
        
        body = ""
        pdf_text = ""
        saved_attachments = []
        
        # Read MIME parts
        if msg.is_multipart():
            for part in msg.walk():
                content_type = part.get_content_type()
                content_disposition = str(part.get("Content-Disposition"))
                filename = part.get_filename()
                
                if filename or "attachment" in content_disposition:
                    if filename:
                        filename = decode_mime_words(filename)
                    else:
                        filename = "unnamed_attachment.bin"
                    
                    # Clean filename to avoid path traversal
                    filename = os.path.basename(filename)
                    payload = part.get_payload(decode=True)
                    
                    if payload:
                        task_dir = os.path.join(secure_attachments_dir, task_id)
                        os.makedirs(task_dir, exist_ok=True)
                        file_path = os.path.join(task_dir, filename)
                        
                        with open(file_path, "wb") as f:
                            f.write(payload)
                        
                        saved_attachments.append(file_path)
                        
                        # If it is a PDF, extract its text content
                        if filename.lower().endswith(".pdf"):
                            pdf_text += f"\n[Documento Anexo: {filename}]\n"
                            pdf_text += extract_pdf_text(file_path)
                
                elif content_type == "text/plain":
                    payload = part.get_payload(decode=True)
                    if payload:
                        charset = part.get_content_charset() or 'utf-8'
                        try:
                            body += payload.decode(charset, errors='replace')
                        except Exception:
                            body += payload.decode('latin1', errors='replace')
        else:
            payload = msg.get_payload(decode=True)
            if payload:
                charset = msg.get_content_charset() or 'utf-8'
                try:
                    body = payload.decode(charset, errors='replace')
                except Exception:
                    body = payload.decode('latin1', errors='replace')
        
        matching_emails.append({
            "task_id": task_id,
            "message_id": message_id,
            "data_ingestao": msg_date.isoformat(),
            "remetente_original": sender,
            "sender_email": sender_email,
            "assunto": subject,
            "body": body,
            "pdf_text": pdf_text,
            "attachments": saved_attachments
        })
        
    print(f"Extracted {len(matching_emails)} emails in the target window.")
    return matching_emails
