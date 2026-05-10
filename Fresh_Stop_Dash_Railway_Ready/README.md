# Fresh Stop Dash

Smart Fuel Pricing Optimization Dashboard

## Features
- Multi-station pricing with strategy buttons (A/B/M) and ML Optimal pricing
- Role-based access (Admin, Manager, Viewer)
- Document management with categories
- SharePoint integration (Device Code Flow)
- Email preview + real Outlook sending
- Full audit trails

## Deployment (Railway)

1. Push this repo to GitHub
2. Connect to Railway
3. Add these environment variables:
   - `SHAREPOINT_CLIENT_ID`
   - `SHAREPOINT_TENANT_ID`
   - `SHAREPOINT_SITE_URL`
   - `SECRET_KEY` (strong random string)

## Local Run
```bash
pip install -r requirements.txt
python app.py
```

## Login
Contact your administrator for credentials. New accounts can only be created by an Admin.
