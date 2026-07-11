import os
import sys
import asyncio
import re
import logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
from datetime import datetime
from telethon import TelegramClient, events
from google.oauth2.service_account import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import pickle
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
from apscheduler.schedulers.asyncio import AsyncIOScheduler
import pytz  
from PIL import Image
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas

if sys.platform == 'win32':
    asyncio.set_event_loop_policy(asyncio.WindowsSelectorEventLoopPolicy())

# --- CONFIGURATIONS ---
API_ID = 38057630                 
API_HASH = '05445482addf82f465d77229b06f696e'       
TARGET_GROUP =-1005188583974  

SPREADSHEET_ID = '1kk2_7H567VTdOXLUIrxc4xAtPrpqUtlwyiDHlGl4lvI'
DRIVE_FOLDER_ID = '1xdm2a1ZwtOuignPt-3hzepIqLyL-0Cel' 

# --- GOOGLE INITIALIZATION ---
# Change from absolute Windows paths (E:\...) to relative paths
SHEET_SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
sheets_creds = Credentials.from_service_account_file('credentials.json', scopes=SHEET_SCOPES)
sheets_service = build('sheets', 'v4', credentials=sheets_creds)

DRIVE_SCOPES = ['https://www.googleapis.com/auth/drive.file']
# --- GOOGLE AUTHENTICATION SYSTEM FOR HEADLESS SERVER ---
from google.oauth2 import service_account
sheets_creds = service_account.Credentials.from_service_account_file(
'credentials.json', scopes=DRIVE_SCOPES)
drive_service = build('drive', 'v3', credentials=sheets_creds)
client = TelegramClient('session_master', API_ID, API_HASH)
sheet_lock = asyncio.Lock()
ALBUM_COOLDOWN = 1.5  
LOCAL_TZ = pytz.timezone('Asia/Phnom_Penh')
# --- PDF GENERATOR UTILITY ---
def create_pdf_from_images(image_paths, pdf_filename):
    """Compiles a list of local image files into a single structured PDF document."""
    try:
        c = canvas.Canvas(pdf_filename, pagesize=letter)
        width, height = letter
        
        for img_path in image_paths:
            try:
                with Image.open(img_path) as img:
                    img_w, img_h = img.size
                    # Scale logic to fit Letter page boundaries cleanly
                    aspect = img_h / float(img_w)
                    if aspect > 1:
                        draw_w = height / aspect if (height / aspect) < width else width
                        draw_h = draw_w * aspect
                    else:
                        draw_w = width
                        draw_h = width * aspect
                        if draw_h > height:
                            draw_h = height
                            draw_w = draw_h / aspect
                            
                    x_offset = (width - draw_w) / 2
                    y_offset = (height - draw_h) / 2
                    c.drawImage(img_path, x_offset, y_offset, width=draw_w, height=draw_h)
                    c.showPage()
            except Exception as e:
                print(f"Skipping corrupt frame image file {img_path}: {e}")
        c.save()
        return True
    except Exception as e:
        print(f"Error compiling structural PDF build: {e}")
        return False

# --- CORE DATA OPERATIONS ---
def upload_to_drive(file_path, filename):
    file_metadata = {'name': filename, 'parents': [DRIVE_FOLDER_ID]}
    media = MediaFileUpload(file_path, resumable=True)
    file = drive_service.files().create(body=file_metadata, media_body=media, fields='id, webViewLink').execute()
    return file.get('webViewLink')

def build_hyperlink_formula(links_list, label_prefix="Link"):
    if not links_list or len(links_list) == 0:
        return "None"
    valid_links = [str(url) for url in links_list if url and str(url).startswith('http')]
    if not valid_links:
        return "None"
    if len(valid_links) == 1:
        return f'=HYPERLINK("{valid_links[0]}", "View File")'
    formula_parts = []
    for idx, url in enumerate(valid_links, start=1):
        formula_parts.append(f'HYPERLINK("{url}", "{label_prefix} #{idx}")')
    return "=" + " & CHAR(10) & ".join(formula_parts)

def append_to_sheet(utc_date, user, text, image_links, pdf_link):
    local_dt = utc_date.astimezone(LOCAL_TZ)
    formatted_time = local_dt.strftime("%Y-%m-%d %H:%M:%S")
    links_val = build_hyperlink_formula(image_links, "Image")
    pdf_val = f'=HYPERLINK("{pdf_link}", "PDF Document")' if pdf_link and pdf_link != "None" else "None"
    
    body = {'values': [[formatted_time, str(user), str(text), links_val, "Pending", "", "", "", "", pdf_val]]}
    sheets_service.spreadsheets().values().append(
        spreadsheetId=SPREADSHEET_ID, range="Sheet1!A:J",
        valueInputOption="USER_ENTERED", body=body
    ).execute()

def update_existing_row_data(search_text, new_text, new_image_links, new_pdf_link):
    result = sheets_service.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID, range="Sheet1!A:J"
    ).execute()
    rows = result.get('values', [])
    
    for idx, row in reversed(list(enumerate(rows))):
        if len(row) >= 3 and search_text in row[2]:
            row_num = idx + 1
            status = row[4].strip().lower()
            
            if status == "pending":
                current_desc = row[2]
                updated_desc = f"{current_desc} | Edit: {new_text}" if new_text else current_desc
                
                current_formula = row[3] if len(row) >= 4 else "None"
                combined_links = re.findall(r'https://[^\s",)]+', current_formula) if "http" in current_formula else []
                combined_links.extend(new_image_links)
                new_formula_val = build_hyperlink_formula(combined_links, "Image")
                
                pdf_val = f'=HYPERLINK("{new_pdf_link}", "PDF Document")' if new_pdf_link and new_pdf_link != "None" else (row[9] if len(row) >= 10 else "None")
                
                sheets_service.spreadsheets().values().update(
                    spreadsheetId=SPREADSHEET_ID, range=f"Sheet1!C{row_num}:D{row_num}",
                    valueInputOption="USER_ENTERED", body={'values': [[updated_desc, new_formula_val]]}
                ).execute()
                
                sheets_service.spreadsheets().values().update(
                    spreadsheetId=SPREADSHEET_ID, range=f"Sheet1!J{row_num}",
                    valueInputOption="USER_ENTERED", body={'values': [[pdf_val]]}
                ).execute()
                return True
    return False

def resolve_existing_row(search_text, resolver_user, resolution_notes, resolution_links_list, pdf_link):
    result = sheets_service.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID, range="Sheet1!A:J"
    ).execute()
    rows = result.get('values', [])
    links_val = build_hyperlink_formula(resolution_links_list, "Image")
    now_local = datetime.now(LOCAL_TZ).strftime("%Y-%m-%d %H:%M:%S")
    
    for idx, row in reversed(list(enumerate(rows))):
        if len(row) >= 3 and search_text in row[2]:
            row_num = idx + 1
            
            # Update Columns E to I (Status, Resolved Time, Resolved By, Resolution Notes, Close Attachment)
            update_body = {
                'values': [["Resolved", now_local, str(resolver_user), str(resolution_notes), links_val]]
            }
            sheets_service.spreadsheets().values().update(
                spreadsheetId=SPREADSHEET_ID, range=f"Sheet1!E{row_num}:I{row_num}",
                valueInputOption="USER_ENTERED", body=update_body
            ).execute()
            
            # Column J is completely skipped here so the original report PDF link is never changed or lost.
            return True
    return False

# --- ALBUM CACHE BACKGROUND WORKER ---
album_cache = {}

async def process_grouped_message(grouped_id):
    await asyncio.sleep(ALBUM_COOLDOWN)
    async with sheet_lock:
        if grouped_id not in album_cache: return
        events_list = album_cache.pop(grouped_id)
        first_event = events_list[0]
        
        combined_text = "".join([e.message.message + " " for e in events_list if e.message.message]).strip()
        sender = await first_event.get_sender()
        user = getattr(sender, 'username', None) or getattr(sender, 'first_name', 'Unknown')
        
        local_paths = []
        uploaded_links = []
        timestamp = datetime.now(LOCAL_TZ).strftime("%Y%m%d_%H%M%S")
        prefix = "UPDATE" if combined_text.lower().startswith('/update') else ("RESOLVED" if combined_text.lower().startswith('/done') else "PROBLEM")
        
        for idx, e in enumerate(events_list):
            if e.message.media:
                path = await e.message.download_media(file='temp_media/')
                if path:
                    local_paths.append(path)
                    filename = f"{prefix}_{timestamp}_{idx+1}_{os.path.basename(path)}"
                    link = upload_to_drive(path, filename)
                    uploaded_links.append(link)
        
        # Compile dynamic PDF report if multiple files exist
        pdf_link = "None"
        if len(local_paths) > 0:
            pdf_filename = f"temp_media/REPORT_{prefix}_{timestamp}.pdf"
            if create_pdf_from_images(local_paths, pdf_filename):
                pdf_link = upload_to_drive(pdf_filename, os.path.basename(pdf_filename))
                os.remove(pdf_filename)
            for p in local_paths:
                if os.path.exists(p): os.remove(p)
        
        if combined_text.lower().startswith('/update') and first_event.message.is_reply:
            reply_msg = await first_event.get_reply_message()
            orig_text = reply_msg.message or ""
            if orig_text:
                success = update_existing_row_data(orig_text, combined_text.replace('/update','',1).strip(), uploaded_links, pdf_link)
                if success: await first_event.respond("📝 Row successfully updated with compiled PDF evidence!")
            return
            
        if combined_text.lower().startswith('/done') and first_event.message.is_reply:
            reply_msg = await first_event.get_reply_message()
            orig_text = reply_msg.message or ""
            if orig_text:
                success = resolve_existing_row(orig_text, user, combined_text.replace('/done','',1).strip(), uploaded_links, pdf_link)
                if success: await first_event.respond("✅ Closed! Row updated with closure PDF summary.")
            return
            
        if combined_text.startswith('/'): return
        try: append_to_sheet(first_event.date, user, combined_text, uploaded_links, pdf_link)
        except Exception as e: print(f"Error: {e}")

# ==========================================
#          MAIN LIVE PIPELINE
# ==========================================
@client.on(events.NewMessage(chats=TARGET_GROUP))
async def pipeline_handler(event):
    print(f"Received message: {event.message.text}")  # <-- Injected line
    if event.message.grouped_id is not None:
        gid = event.message.grouped_id
        if gid not in album_cache:
            album_cache[gid] = [event]
            asyncio.create_task(process_grouped_message(gid))
        else:
            album_cache[gid].append(event)
        return

    msg_text = event.message.message or ""
    sender = await event.get_sender()
    user = getattr(sender, 'username', None) or getattr(sender, 'first_name', 'Unknown')
    
    # Process Single Item Entries (Download and compile immediately)
    uploaded_links = []
    pdf_link = "None"
    
    if event.message.media and not msg_text.strip().startswith('/'):
        path = await event.message.download_media(file='temp_media/')
        if path:
            timestamp = datetime.now(LOCAL_TZ).strftime("%Y%m%d_%H%M%S")
            link = upload_to_drive(path, f"PROBLEM_{timestamp}_{os.path.basename(path)}")
            uploaded_links.append(link)
            
            pdf_filename = f"temp_media/REPORT_PROBLEM_{timestamp}.pdf"
            if create_pdf_from_images([path], pdf_filename):
                pdf_link = upload_to_drive(pdf_filename, os.path.basename(pdf_filename))
                os.remove(pdf_filename)
            os.remove(path)

    # Command Router Layer
    if msg_text.strip().startswith('/'):
        # 1. Single-image / Single-text Update Router
        if msg_text.strip().lower().startswith('/update') and event.message.is_reply:
            reply_msg = await event.get_reply_message()
            orig_text = reply_msg.message or ""
            if orig_text:
                if event.message.media:
                    path = await event.message.download_media(file='temp_media/')
                    if path:
                        timestamp = datetime.now(LOCAL_TZ).strftime("%Y%m%d_%H%M%S")
                        link = upload_to_drive(path, f"UPDATE_{timestamp}_{os.path.basename(path)}")
                        uploaded_links.append(link)
                        
                        pdf_filename = f"temp_media/REPORT_UPDATE_{timestamp}.pdf"
                        if create_pdf_from_images([path], pdf_filename):
                            pdf_link = upload_to_drive(pdf_filename, os.path.basename(pdf_filename))
                            os.remove(pdf_filename)
                        os.remove(path)
                        
                async with sheet_lock:
                    update_existing_row_data(orig_text, msg_text.replace('/update','',1).strip(), uploaded_links, pdf_link)
                    await event.respond("📝 Entry updated.")
                    
        # 2. Single-image / Single-text Resolution Router (FIXED)
        elif msg_text.strip().lower().startswith('/done') and event.message.is_reply:
            reply_msg = await event.get_reply_message()
            orig_text = reply_msg.message or ""
            if orig_text:
                if event.message.media:
                    path = await event.message.download_media(file='temp_media/')
                    if path:
                        timestamp = datetime.now(LOCAL_TZ).strftime("%Y%m%d_%H%M%S")
                        link = upload_to_drive(path, f"RESOLVED_{timestamp}_{os.path.basename(path)}")
                        uploaded_links.append(link)
                        
                        # Generate the PDF document tracking metric for resolution phase
                        pdf_filename = f"temp_media/REPORT_RESOLVED_{timestamp}.pdf"
                        if create_pdf_from_images([path], pdf_filename):
                            pdf_link = upload_to_drive(pdf_filename, os.path.basename(pdf_filename))
                            os.remove(pdf_filename)
                        os.remove(path)
                        
                async with sheet_lock:
                    resolve_existing_row(orig_text, user, msg_text.replace('/done','',1).strip(), uploaded_links, pdf_link)
                    await event.respond("✅ Task closed.")
        return

    # Standard Log Entry Fallthrough
    async with sheet_lock:
        try: append_to_sheet(event.date, user, msg_text, uploaded_links, pdf_link)
        except Exception as e: print(f"Error: {e}")
async def main():
    print("Starting Telegram Client...")
    await client.start(bot_token='8327734059:AAGzhXIdaS05Lm1faveX-GipV2wrTvRFd3M')
    print("----------------------------------------")
    print("PDF ENGINE OPERATIONAL: Listening for Telegram messages...")
    print("----------------------------------------")

    # This blocks the async function from completing, keeping the cloud server awake
    await client.run_until_disconnected()

if __name__ == '__main__':
    import asyncio
    asyncio.run(main())
