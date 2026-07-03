import os
import json
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload

# Shared Google Drive Folder ID
DEFAULT_FOLDER_ID = "1cDNkx6L806mgws_CbPve-ZfSKfvEwhdk"

def get_drive_service():
    """Initializes and returns the Google Drive API service using Service Account credentials."""
    creds_json = os.environ.get("GDRIVE_SERVICE_ACCOUNT_JSON")
    
    if creds_json:
        # Load from GitHub Secrets environment variable
        try:
            info = json.loads(creds_json)
            if isinstance(info, dict) and "client_email" in info:
                print(f"[GDrive] Authenticating using Service Account: {info['client_email']}")
                print(f"[GDrive] IMPORTANT: Please make sure this email is added as an 'Editor' to the Google Drive Folder (ID: {DEFAULT_FOLDER_ID})")
            creds = service_account.Credentials.from_service_account_info(
                info, scopes=["https://www.googleapis.com/auth/drive.file", "https://www.googleapis.com/auth/drive"]
            )
            return build("drive", "v3", credentials=creds)
        except Exception as e:
            print(f"[GDrive] Error parsing GDRIVE_SERVICE_ACCOUNT_JSON env var: {e}")
            
    # Fallback to local credentials.json
    local_creds_path = os.path.join(os.path.dirname(__file__), "credentials.json")
    if os.path.exists(local_creds_path):
        try:
            with open(local_creds_path, "r", encoding="utf-8") as f:
                info = json.load(f)
            if isinstance(info, dict) and "client_email" in info:
                print(f"[GDrive] Authenticating using local credentials: {info['client_email']}")
                print(f"[GDrive] IMPORTANT: Please make sure this email is added as an 'Editor' to the Google Drive Folder (ID: {DEFAULT_FOLDER_ID})")
            creds = service_account.Credentials.from_service_account_file(
                local_creds_path, scopes=["https://www.googleapis.com/auth/drive.file", "https://www.googleapis.com/auth/drive"]
            )
            return build("drive", "v3", credentials=creds)
        except Exception as e:
            print(f"[GDrive] Error loading local credentials.json: {e}")
            
    print("[GDrive] Warning: No valid Google Drive credentials found. Skipping upload.")
    return None


def get_or_create_folder(service, folder_name, parent_id):
    """Checks if a folder with the given name exists under parent_id, otherwise creates it.
    Nếu tìm thấy NHIỀU folder cùng tên (dấu hiệu đã từng bị tạo trùng do bug/race condition trước
    đây), luôn chọn folder có createdTime SỚM NHẤT một cách nhất quán (thay vì phần tử đầu tiên
    Drive API trả về, thứ tự không đảm bảo) để không tạo thêm bản sao/không nhảy lung tung giữa các
    bản trùng qua từng lần chạy — đồng thời cảnh báo rõ để người dùng biết mà dọn thủ công trên Drive."""
    if not service:
        return parent_id
    try:
        query = f"name = '{folder_name}' and mimeType = 'application/vnd.google-apps.folder' and '{parent_id}' in parents and trashed = false"
        results = service.files().list(q=query, fields="files(id, createdTime)").execute()
        files = results.get("files", [])
        if files:
            if len(files) > 1:
                print(f"[GDrive] CẢNH BÁO: có {len(files)} folder trùng tên '{folder_name}' trong parent {parent_id} "
                      f"(ids: {[f['id'] for f in files]}) - đang dùng folder tạo sớm nhất, nên dọn thủ công trên Drive.")
            files.sort(key=lambda f: f.get("createdTime", ""))
            chosen_id = files[0]["id"]
            print(f"[GDrive] Resolved folder '{folder_name}' (parent {parent_id}) -> {chosen_id}")
            return chosen_id
        else:
            folder_metadata = {
                'name': folder_name,
                'mimeType': 'application/vnd.google-apps.folder',
                'parents': [parent_id]
            }
            folder = service.files().create(body=folder_metadata, fields='id').execute()
            new_id = folder.get('id')
            print(f"[GDrive] Created new folder '{folder_name}' (parent {parent_id}) -> {new_id}")
            return new_id
    except Exception as e:
        print(f"[GDrive] Error getting or creating folder '{folder_name}': {e}")
        return parent_id


def upload_file(file_path, folder_id=None, sector=None, ticker=None):
    """
    Uploads a file to a specific Google Drive folder.
    Creates subfolders based on sector and ticker if they are supplied.
    Returns: (file_id, web_view_link) or (None, None)
    """
    if not os.path.exists(file_path):
        print(f"[GDrive] File not found: {file_path}")
        return None, None

    BANK_FOLDER_ID = "1WZnuR6MH2914b-efJj1_4XQel2ZaGRKw"
    
    is_bank = False
    if sector:
        sec_lower = sector.lower()
        if any(k in sec_lower for k in ["bank", "ngân hàng", "tài chính", "financial"]):
            is_bank = True
            
    base_folder_id = folder_id if folder_id is not None else (BANK_FOLDER_ID if is_bank else DEFAULT_FOLDER_ID)

    service = get_drive_service()

    # Cấp "Ngành" (sector_level_folder_id): Bank -> chính BANK_FOLDER_ID (không cần thêm cấp con);
    # ngành khác -> tạo/tìm subfolder theo tên sector bên dưới base_folder_id.
    sector_level_folder_id = base_folder_id
    if service and not is_bank and sector:
        sector_level_folder_id = get_or_create_folder(service, sector, base_folder_id)
    elif not service:
        print(f"[GDrive] No service account available to resolve subfolders. Using base folder ID: {base_folder_id}")

    # Cấp "Mã" (resolved_folder_id, đủ 2 cấp Ngành/Mã — đúng cấu trúc Ngành/Mã/file mong muốn):
    # nếu folder Mã đã tồn tại trong folder Ngành thì get_or_create_folder trả về ID có sẵn (chỉ add
    # file vào); nếu folder Ngành chưa có thì bước ở trên đã tạo Ngành trước, rồi mới tạo Mã ở đây.
    resolved_folder_id = sector_level_folder_id
    if service and ticker:
        resolved_folder_id = get_or_create_folder(service, ticker.upper(), sector_level_folder_id)

    # ── Try Web App Upload if URL is configured ──
    webapp_url = os.environ.get("GDRIVE_WEBAPP_URL")
    if webapp_url:
        import base64
        import requests
        
        file_name = os.path.basename(file_path)
        actual_ticker = ticker.upper() if ticker else os.path.basename(os.path.dirname(file_path)).upper()
        print(f"[GDrive] Uploading via Google Apps Script Web App: {file_name} (Ticker: {actual_ticker})...")
        
        mime_type = "application/octet-stream"
        if file_name.endswith(".pdf"):
            mime_type = "application/pdf"
        elif file_name.endswith(".xlsx"):
            mime_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        elif file_name.endswith(".json"):
            mime_type = "application/json"
            
        try:
            with open(file_path, "rb") as f:
                content = f.read()
            encoded_content = base64.b64encode(content).decode("utf-8")
            
            # folderId ở đây đã là ID cấp "Mã" (Ngành/Mã đã resolve đủ 2 cấp ở trên) — KHÔNG gửi kèm
            # "ticker" nữa (2026-07, phát hiện bug folder lồng "Ngành/Mã/Mã/"): nghi vấn cao là Apps
            # Script phía Google (không nằm trong repo này) tự tạo THÊM 1 cấp folder theo ticker khi
            # nhận field này, dù Python đã resolve folder đích chính xác rồi. Nếu Apps Script vẫn cần
            # ticker cho mục đích khác (đặt tên file, log...), thêm lại field riêng KHÔNG dùng để tạo
            # folder, hoặc sửa thẳng trong Apps Script để bỏ bước tự tạo folder theo ticker.
            payload = {
                "token": "FA_PIPELINE_SECRET_2026",
                "folderId": resolved_folder_id,
                "fileName": file_name,
                "fileContent": encoded_content,
                "mimeType": mime_type
            }
            
            # Send POST request to Google Apps Script Web App
            resp = requests.post(webapp_url, json=payload, timeout=90)
            resp.raise_for_status()
            result = resp.json()
            
            if result.get("status") == "success":
                print(f"[GDrive] Uploaded successfully via Web App: {file_name} -> {result.get('webViewLink')}")
                return result.get("fileId"), result.get("webViewLink")
            else:
                print(f"[GDrive] Web App upload failed: {result.get('message')}. Trying Service Account fallback...")
        except Exception as err:
            print(f"[GDrive] Error during Web App upload: {err}. Trying Service Account fallback...")
            
    if not service:
        return None, None
        
    file_name = os.path.basename(file_path)
    
    # Determine mimeType based on file extension
    mime_type = "application/octet-stream"
    if file_name.endswith(".pdf"):
        mime_type = "application/pdf"
    elif file_name.endswith(".xlsx"):
        mime_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    elif file_name.endswith(".json"):
        mime_type = "application/json"
        
    file_metadata = {
        "name": file_name,
        "parents": [resolved_folder_id]
    }
    
    media = MediaFileUpload(file_path, mimetype=mime_type, resumable=True)
    
    try:
        print(f"[GDrive] Uploading {file_name} to folder {resolved_folder_id}...")
        
        # Check if file with same name already exists in the destination folder to avoid duplicates (optional but neat)
        query = f"name = '{file_name}' and '{resolved_folder_id}' in parents and trashed = false"
        results = service.files().list(q=query, fields="files(id)").execute()
        files = results.get("files", [])
        
        if files:
            # Update existing file
            existing_file_id = files[0]["id"]
            print(f"[GDrive] Updating existing file (ID: {existing_file_id})...")
            file = service.files().update(
                fileId=existing_file_id,
                media_body=media,
                fields="id, webViewLink"
            ).execute()
        else:
            # Create new file
            file = service.files().create(
                body=file_metadata,
                media_body=media,
                fields="id, webViewLink"
            ).execute()
            
        # Set permission to anyone with link can view (publicly readable so that dashboard could link to them)
        try:
            permission = {
                'type': 'anyone',
                'role': 'reader',
            }
            service.permissions().create(
                fileId=file.get("id"),
                body=permission
            ).execute()
        except Exception as perm_err:
            # If sharing permission fails, just log it and proceed
            print(f"[GDrive] Note: Could not set public permission: {perm_err}")

        print(f"[GDrive] Uploaded successfully: {file_name} -> {file.get('webViewLink')}")
        return file.get("id"), file.get("webViewLink")
        
    except Exception as e:
        print(f"[GDrive] Upload failed for {file_name}: {e}")
        return None, None
