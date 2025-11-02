import imaplib
import email
import os
import re
import time
import socket
from bs4 import BeautifulSoup
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# ----------------------------
# Email credentials (hardcoded)
# ----------------------------
EMAIL_USER = "sarthakmohite094@gmail.com"
EMAIL_PASS = "ntfv uywk mxef brum"

# ----------------------------
# Google Sheets setup
# ----------------------------
scope = ["https://spreadsheets.google.com/feeds",
         "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name("credentials.json", scope)
client = gspread.authorize(creds)

sheet = client.open("Student_data").worksheet("checked_strings")

def get_next_serial():
    """Get next serial number based on last row."""
    records = sheet.get_all_values()
    if len(records) <= 1:
        return 1
    return len(records)

def save_to_sheet(name, phone, email_addr, event_name):
    serial_no = get_next_serial()
    sheet.append_row([serial_no, name, phone, email_addr, event_name])
    print(f"✅ Saved: {name} | Event: {event_name} | Serial No: {serial_no}")

# ----------------------------
# Email fetching with retry logic
# ----------------------------
def fetch_emails(max_retries=3, retry_delay=5):
    """Fetch emails with retry logic and better error handling."""
    
    for attempt in range(max_retries):
        mail = None
        try:
            # Set socket timeout
            socket.setdefaulttimeout(30)
            
            # Connect to Gmail
            print(f"🔌 Connecting to Gmail (attempt {attempt + 1}/{max_retries})...")
            mail = imaplib.IMAP4_SSL("imap.gmail.com", 993)
            
            # Login
            mail.login(EMAIL_USER, EMAIL_PASS)
            print("✅ Successfully logged in")
            
            # Select inbox
            mail.select("inbox")
            
            # Search for unread emails
            status, data = mail.search(None, '(UNSEEN SUBJECT "New Event Registration")')
            
            if status != 'OK':
                print("⚠️ Search failed")
                return []
            
            email_ids = data[0].split()
            messages = []
            
            print(f"📬 Found {len(email_ids)} unread registration emails")
            
            for e_id in email_ids:
                try:
                    _, msg_data = mail.fetch(e_id, "(RFC822)")
                    raw_msg = msg_data[0][1]
                    msg = email.message_from_bytes(raw_msg)
                    
                    body = ""
                    if msg.is_multipart():
                        for part in msg.walk():
                            if part.get_content_type() in ["text/plain", "text/html"]:
                                try:
                                    body = part.get_payload(decode=True).decode('utf-8', errors='ignore')
                                    break
                                except:
                                    continue
                    else:
                        try:
                            body = msg.get_payload(decode=True).decode('utf-8', errors='ignore')
                        except:
                            body = str(msg.get_payload())
                    
                    messages.append({
                        "subject": msg["subject"], 
                        "from": msg["from"], 
                        "text": body
                    })
                except Exception as e:
                    print(f"⚠️ Error processing email {e_id}: {e}")
                    continue
            
            # Close and logout
            mail.close()
            mail.logout()
            
            return messages
            
        except imaplib.IMAP4.abort as e:
            print(f"❌ IMAP connection aborted: {e}")
            if mail:
                try:
                    mail.logout()
                except:
                    pass
            
            if attempt < max_retries - 1:
                print(f"⏳ Retrying in {retry_delay} seconds...")
                time.sleep(retry_delay)
            else:
                print("❌ Max retries reached. Skipping this cycle.")
                return []
                
        except socket.error as e:
            print(f"❌ Socket error: {e}")
            if mail:
                try:
                    mail.logout()
                except:
                    pass
            
            if attempt < max_retries - 1:
                print(f"⏳ Retrying in {retry_delay} seconds...")
                time.sleep(retry_delay)
            else:
                print("❌ Max retries reached. Skipping this cycle.")
                return []
                
        except Exception as e:
            print(f"❌ Unexpected error: {e}")
            if mail:
                try:
                    mail.logout()
                except:
                    pass
            
            if attempt < max_retries - 1:
                print(f"⏳ Retrying in {retry_delay} seconds...")
                time.sleep(retry_delay)
            else:
                print("❌ Max retries reached. Skipping this cycle.")
                return []
    
    return []

# ----------------------------
# Extract registration details
# ----------------------------
def extract_details(email_text):
    soup = BeautifulSoup(email_text, "html.parser")
    clean_text = soup.get_text(separator="\n")
    
    # Print full cleaned text for debugging
    print(f"📝 CLEANED EMAIL TEXT:\n{clean_text[:500]}...\n{'='*80}\n")
    
    # Initialize variables
    name = ""
    phone = ""
    email_addr = ""
    event = ""
    
    # Split text into lines for easier parsing
    lines = [line.strip() for line in clean_text.split('\n') if line.strip()]
    
    # Try to parse line by line
    for i, line in enumerate(lines):
        line_lower = line.lower()
        
        # Check for Name
        if not name and any(keyword in line_lower for keyword in ['name:', 'full name:', 'student name:']):
            parts = line.split(':', 1)
            if len(parts) > 1:
                name = parts[1].strip()
            elif i + 1 < len(lines):
                name = lines[i + 1].strip()
        
        # Check for Phone
        if not phone and any(keyword in line_lower for keyword in ['phone:', 'mobile:', 'contact:', 'phone number:']):
            parts = line.split(':', 1)
            if len(parts) > 1:
                phone = parts[1].strip()
            elif i + 1 < len(lines):
                phone = lines[i + 1].strip()
        
        # Check for Email
        if not email_addr and any(keyword in line_lower for keyword in ['email:', 'e-mail:', 'email address:']):
            parts = line.split(':', 1)
            if len(parts) > 1:
                email_addr = parts[1].strip()
            elif i + 1 < len(lines):
                email_addr = lines[i + 1].strip()
        
        # Check for Event
        if not event and any(keyword in line_lower for keyword in [
            'event:', 'event name:', 'registered for:', 'registration for:', 
            'event title:', 'workshop:', 'program:', 'course:'
        ]):
            parts = line.split(':', 1)
            if len(parts) > 1 and parts[1].strip():
                event = parts[1].strip()
            elif i + 1 < len(lines):
                potential_event = lines[i + 1].strip()
                if not any(field in potential_event.lower() for field in ['name:', 'phone:', 'email:', 'contact:']):
                    event = potential_event
    
    # Fallback to regex if needed
    if not name:
        name_match = re.search(r"(?:Name|Full Name|Student Name)[:\-\s]*(.+?)(?=\n|$)", clean_text, re.IGNORECASE)
        name = name_match.group(1).strip() if name_match else ""
    
    if not phone:
        phone_match = re.search(r"(?:Phone|Phone Number|Mobile|Contact)[:\-\s]*(.+?)(?=\n|$)", clean_text, re.IGNORECASE)
        phone = phone_match.group(1).strip() if phone_match else ""
    
    if not email_addr:
        email_match = re.search(r"(?:Email|Email Address|E-mail)[:\-\s]*([^\s\n]+@[^\s\n]+)", clean_text, re.IGNORECASE)
        email_addr = email_match.group(1).strip() if email_match else ""
    
    if not event:
        event_match = re.search(r"(?:Event|Event Name|Registered for)[:\-\s]+(.+?)(?=\n(?:Name|Phone|Email)|$)", clean_text, re.IGNORECASE | re.DOTALL)
        event = event_match.group(1).strip() if event_match else ""
    
    # Clean up
    name = re.sub(r'\s+', ' ', name).strip()
    phone = re.sub(r'[^\d+\-() ]', '', phone).strip()
    email_addr = email_addr.strip()
    
    if event:
        event = re.sub(r'\s+', ' ', event).strip()
        event = event.split('\n')[0].strip()
        event = re.sub(r'\s*(?:registration|form)$', '', event, flags=re.IGNORECASE).strip()
    
    print(f"  📋 Extracted: Name='{name}', Phone='{phone}', Email='{email_addr}', Event='{event}'\n")
    
    return {
        "Name": name,
        "Phone": phone,
        "Email": email_addr,
        "Event": event
    }

# ----------------------------
# Main processing loop
# ----------------------------
if __name__ == "__main__":
    print("🤖 Starting continuous email processing agent...")
    
    if not EMAIL_USER or not EMAIL_PASS:
        print("❌ EMAIL_USER or EMAIL_PASS not set")
        exit(1)
    
    consecutive_errors = 0
    max_consecutive_errors = 5
    
    try:
        while True:
            try:
                emails = fetch_emails()
                consecutive_errors = 0  # Reset on success
                
                if emails:
                    print(f"📨 Processing {len(emails)} registration emails")
                    
                for mail in emails:
                    print(f"📧 Processing: {mail['subject']}")
                    info = extract_details(mail["text"])
                    
                    if info and info["Name"]:
                        save_to_sheet(info["Name"], info["Phone"], info["Email"], info["Event"])
                    else:
                        print(f"⚠️ Could not parse email. Extracted: {info}\n")
                
                print("🕒 Waiting 30 seconds before next check...\n")
                time.sleep(30)
                
            except Exception as e:
                consecutive_errors += 1
                print(f"\n❌ Error in processing cycle: {e}")
                
                if consecutive_errors >= max_consecutive_errors:
                    print(f"❌ Too many consecutive errors ({consecutive_errors}). Exiting.")
                    break
                
                print(f"⏳ Waiting 60 seconds before retry...\n")
                time.sleep(60)
    
    except KeyboardInterrupt:
        print("\n🛑 Agent terminated by user.")
    except Exception as e:
        print(f"\n❌ Fatal error: {e}")
        import traceback
        traceback.print_exc()