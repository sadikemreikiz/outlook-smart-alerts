# Outlook Smart Alerts 

An Outlook Add-in that intercepts the **Send** button and shows a confirmation dialog before the email is sent.

## What It Does

When the user clicks **Send**, a popup appears with 3 questions:

1. Does this email contain confidential information? (Yes / No)
2. Priority level? (Low / Normal / High)
3. Do you confirm sending? (Send / Cancel)

If the user clicks **Send** → email is sent.  
If the user clicks **Cancel** → email is blocked.

## Live Dialog Preview

👉 [Click here to preview the dialog](https://sadikemreikiz.github.io/outlook-smart-alerts/dialog.html)

## File Structure

```
outlook-smart-alerts/
├── manifest.xml      # Add-in manifest — sideload this into Outlook
├── commands.html     # Event-based runtime page (no UI)
├── commands.js       # OnMessageSend event handler
├── dialog.html       # Confirmation popup UI
└── dialog.js         # Dialog logic (Send / Cancel)
```

## How It Works

```
User clicks Send
    └── OnMessageSend event fires
    └── commands.js opens dialog.html as a popup
    └── User answers the 3 questions
    └── "Send"   → email is sent      (allowEvent: true)
    └── "Cancel" → email is blocked   (allowEvent: false)
```

## How to Test

### Requirements
- Microsoft 365 work or school account (not personal @outlook.com)
- Access to Outlook on the Web: https://outlook.office.com

### Steps

**1. Open Outlook on the Web**
```
https://outlook.office.com
```

**2. Go to Add-ins settings**
```
Settings → View all Outlook settings → Mail → Customize actions → Add-ins
```
Or navigate directly:
```
https://outlook.office.com/mail/options/general/manageAddIns
```

**3. Sideload the manifest**
- Click **"My add-ins"** tab
- Click **"Add a custom add-in"** → **"Add from file"**
- Select `manifest.xml` from this repository

**4. Test**
- Compose a new email
- Click **Send**
- The confirmation dialog should appear

## Tech Stack

- JavaScript (vanilla)
- Office.js (Microsoft Office JavaScript API)
- HTML / CSS
- Manifest: XML (add-in only)
- Hosting: GitHub Pages (HTTPS)

## Requirements

- Office.js event: `OnMessageSend`
- Mailbox requirement set: `1.13`
- Supported: New Outlook, Outlook on the Web
- Not supported: Classic Outlook (VSTO)
