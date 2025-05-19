# Credentials Setup

This directory is for storing your Google API credentials securely for the MATMAX WELLNESS financial model export script.

## How to Set Up Credentials

1. Create a Google Cloud Project
2. Enable the Google Sheets API
3. Create a Service Account
4. Generate a JSON key for the Service Account
5. Save the key as `google_service_account.json` in this directory

The export script will automatically detect and use this file.

## Alternative Method

Instead of storing the credentials file, you can set the `GOOGLE_CREDS_JSON` environment variable with the contents of your service account JSON.

Example:
```bash
export GOOGLE_CREDS_JSON=$(cat /path/to/your-service-account.json)
```

## Security Notes

- Never commit credential files to Git
- Keep your service account keys secure
- The main export script is designed to handle credentials securely
- Temporary credential files are automatically deleted after use 