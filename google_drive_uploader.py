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


def upload_file(file_path, folder_id=DEFAULT_FOLDER_ID):
    """
    Uploads a file to a specific Google Drive folder.
    Returns: (file_id, web_view_link) or (None, None)
    """
    if not os.path.exists(file_path):
        print(f"[GDrive] File not found: {file_path}")
        return None, None
        
    # ── Try Web App Upload if URL is configured ──
    webapp_url = os.environ.get("GDRIVE_WEBAPP_URL")
    if webapp_url:
        import base64
        import requests
        
        file_name = os.path.basename(file_path)
        print(f"[GDrive] Uploading via Google Apps Script Web App: {file_name}...")
        
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
            
            payload = {
                "token": "FA_PIPELINE_SECRET_2026",
                "folderId": folder_id,
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
            
    service = get_drive_service()
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
        "parents": [folder_id]
    }
    
    media = MediaFileUpload(file_path, mimetype=mime_type, resumable=True)
    
    try:
        print(f"[GDrive] Uploading {file_name} to folder {folder_id}...")
        
        # Check if file with same name already exists in the destination folder to avoid duplicates (optional but neat)
        query = f"name = '{file_name}' and '{folder_id}' in parents and trashed = false"
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
