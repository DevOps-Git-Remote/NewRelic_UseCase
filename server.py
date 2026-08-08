import json
import requests
import msal
import logging
from mcp.server.fastmcp import FastMCP

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Initialize the FastMCP server
mcp = FastMCP("M365-Incident-Notifier")

@mcp.tool()
def notify_team_via_m365_graph(eprid: str, email_body: str) -> str:
    """
    Automates an incident triage alert for a given EPRID.
    Fetches dynamic routing config from SharePoint Excel and dispatches 
    alerts via Outlook (Email) and Power Automate Webhooks (Teams).
    """
    tenant_id = "71ced5a9-db6d-473d-a0f5-ef0d7c4f6970"
    client_id = "43f61d12-2f59-4bba-9261-efa0be982e89"
    client_secret = "JFV8Q~mwFh-4Wv.jtoQ1Jr6TWibEmvsQRkkWXaMe" 
    
    sender_email = "ServiceNow-Learnings@w422w.onmicrosoft.com"
    
    # 1. Authenticate 
    try:
        authority = f"https://login.microsoftonline.com/{tenant_id}"
        app = msal.ConfidentialClientApplication(client_id, authority=authority, client_credential=client_secret)
        token_response = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
        
        if "access_token" not in token_response:
            return json.dumps({"status": "ERROR", "message": "Authentication failed with Microsoft Graph."})
    except Exception as e:
        return json.dumps({"status": "ERROR", "message": f"MSAL Auth Error: {str(e)}"})

    headers = {
        "Authorization": f"Bearer {token_response['access_token']}",
        "Content-Type": "application/json"
    }

    # 2. Fetch Dynamic Data from SharePoint Excel
    excel_fields = get_excel_routing_config(eprid, headers)
    
    if excel_fields and "_error" in excel_fields:
        return json.dumps({
            "status": "ERROR",
            "message": f"SHAREPOINT FETCH FAILED: {excel_fields['_error']}"
        })

    to_emails_str = excel_fields.get("To")
    cc_emails_str = excel_fields.get("CC", "")
    channel_workflow = excel_fields.get("Channel", "Unknown_Workflow")

    # 3. Construct Email Payload
    email_subject = f"🚨 Automated Triage Alert: Incident Report for {eprid}"

    email_payload = {
        "message": {
            "subject": email_subject,
            "body": {"contentType": "Text", "content": email_body},
            "toRecipients": format_recipients(to_emails_str),
        },
        "saveToSentItems": "false"
    }
    
    cc_recipients = format_recipients(cc_emails_str)
    if cc_recipients:
        email_payload["message"]["ccRecipients"] = cc_recipients

    # 4. Dispatch Email
    email_status = "Skipped"
    send_mail_url = f"https://graph.microsoft.com/v1.0/users/{sender_email}/sendMail"
    
    try:
        response = requests.post(send_mail_url, headers=headers, json=email_payload)
        response.raise_for_status()
        email_status = "SUCCESS"
    except Exception as e:
        error_details = e.response.text if hasattr(e, 'response') and e.response else str(e)
        logger.error(f"Failed to send email: {error_details}")
        email_status = f"FAILED: {error_details}"
        
    # 5. Dispatch to Teams via Power Automate Webhook
    teams_status = "Skipped (No valid webhook URL found)"
    if channel_workflow.startswith("http"):
        teams_payload = {
            "type": "AdaptiveCard",
            "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
            "version": "1.4",
            "body": [
                {
                    "type": "TextBlock",
                    "text": email_subject,
                    "weight": "Bolder",
                    "size": "Large",
                    "wrap": True
                },
                {
                    "type": "TextBlock",
                    "text": email_body,
                    "wrap": True
                }
            ]
        }
        
        try:
            webhook_response = requests.post(channel_workflow, json=teams_payload, timeout=10)
            webhook_response.raise_for_status()
            teams_status = "SUCCESS"
            logger.info(f"Successfully triggered Teams webhook for EPRID {eprid}.")
        except Exception as e:
            logger.error(f"Failed to post to Teams webhook: {str(e)}")
            teams_status = f"FAILED: {str(e)}"

    # 6. Return Final Delivery Report
    if "FAILED" in email_status and "FAILED" in teams_status:
        return json.dumps({"status": "ERROR", "message": f"Both Email and Teams dispatches failed. Email: {email_status} | Teams: {teams_status}"})

    return json.dumps({
        "status": "SUCCESS",
        "message": f"Dispatch complete for EPRID {eprid}. Email: {email_status} | Teams: {teams_status}",
        "delivery_details": {
            "to": to_emails_str,
            "cc": cc_emails_str,
            "teams_webhook_triggered": True if teams_status == "SUCCESS" else False
        }
    })
    
def get_excel_routing_config(eprid: str, headers: dict, sheet_name: str = "Sheet1") -> dict:
    """Safely resolves the Site ID first, then reads the Excel file."""
    try:
        site_url = "https://graph.microsoft.com/v1.0/sites/w422w.sharepoint.com:/sites/DXC-AI-Core"
        site_resp = requests.get(site_url, headers=headers)
        
        if not site_resp.ok:
            return {"_error": f"Failed to resolve SharePoint Site: {site_resp.text}"}
            
        site_id = site_resp.json().get("id")

        excel_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root:/team_routing.xlsx:/workbook/worksheets('{sheet_name}')/usedRange"
        
        response = requests.get(excel_url, headers=headers)
        if not response.ok:
            return {"_error": f"Failed to read Excel file: HTTP {response.status_code} - {response.text}"}
            
        data = response.json()
        values = data.get('text', [])
        
        if not values or len(values) < 2:
            return {"_error": "Excel file is empty or only contains headers."}
            
        headers_row = values[0]
        
        try:
            eprid_idx = headers_row.index("EPRID")
            to_idx = headers_row.index("To")
            cc_idx = headers_row.index("CC")
            channel_idx = headers_row.index("Channel")
        except ValueError as e:
            return {"_error": f"Missing a required column. Found: {headers_row}"}

        eprid_clean = str(eprid).split('.')[0].strip()

        for row in values[1:]:
            row_eprid = str(row[eprid_idx]).split('.')[0].strip() if len(row) > eprid_idx else ""
            if row_eprid == eprid_clean:
                return {
                    "To": str(row[to_idx]).strip() if len(row) > to_idx else "",
                    "CC": str(row[cc_idx]).strip() if len(row) > cc_idx else "",
                    "Channel": str(row[channel_idx]).strip() if len(row) > channel_idx else "Unknown_Workflow"
                }
                
        return {"_error": f"EPRID {eprid_clean} not found in the SharePoint Excel file."}
        
    except Exception as e:
        return {"_error": f"Exception during Excel fetch: {str(e)}"}

def format_recipients(email_string: str) -> list:
    if not email_string:
        return []
    raw_emails = email_string.replace(',', ';').split(';')
    return [{"emailAddress": {"address": email.strip()}} for email in raw_emails if email.strip()]

# Expose the underlying FastAPI app for Render / Uvicorn deployment
# Expose the underlying FastAPI app for Render / Uvicorn deployment
app = mcp.app

if __name__ == "__main__":
    mcp.run()
