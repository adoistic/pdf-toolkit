# Firebase Setup Guide for PDF Toolkit Licensing

## 1. Create a Firebase Project

1. Go to [Firebase Console](https://console.firebase.google.com/)
2. Click **Add project**
3. Name it (e.g., `pdf-toolkit`)
4. Disable Google Analytics (not needed)
5. Click **Create project**

## 2. Enable Cloud Firestore

1. In the Firebase Console, go to **Build → Firestore Database**
2. Click **Create database**
3. Choose **Start in production mode** (we'll set rules next)
4. Select a region close to your users (e.g., `us-east1` or `europe-west1`)
5. Click **Enable**

## 3. Deploy Security Rules

1. In Firestore, go to the **Rules** tab
2. Replace the default rules with the contents of `firestore.rules` from this project
3. Click **Publish**

## 4. Get Your API Key

1. Go to **Project Settings** (gear icon) → **General**
2. Scroll down to **Your apps**
3. Click **Add app** → choose **Web** (</> icon)
4. Name it (e.g., `pdf-toolkit-client`)
5. Copy the `apiKey` value from the config snippet
6. Also note the `projectId`

## 5. Update license.py

Open `license.py` and replace the placeholder values:

```python
FIREBASE_PROJECT_ID = "pdf-toolkit-abc123"    # Your actual project ID
FIREBASE_API_KEY    = "AIzaSy..."              # Your actual API key
```

## 6. Create a Service Account (for Admin CLI)

1. Go to **Project Settings → Service Accounts**
2. Click **Generate new private key**
3. Save the downloaded JSON file as `service_account.json` in the project directory
4. **IMPORTANT:** Never include this file in the EXE or share it with customers

## 7. Install Admin CLI Dependencies

On your admin machine only:

```bash
pip install firebase-admin
```

## 8. Test the Setup

### Generate a license key:
```bash
python admin_license.py generate --days 365 --note "Test user"
```

### List all licenses:
```bash
python admin_license.py list
```

### Test in the app:
1. Run `python app.py`
2. The license modal should appear
3. Enter the generated key
4. The app should activate and load normally

### Test revocation:
```bash
python admin_license.py revoke PDFT-XXXX-XXXX-XXXX
```
Then restart the app — it should show "License has been revoked" on the next online check.

### Test unbinding (move to a different machine):
```bash
python admin_license.py unbind PDFT-XXXX-XXXX-XXXX
```

## Security Notes

- The **API key** is safe to embed in the app. It only identifies the Firebase project; all access control is enforced by Firestore security rules.
- The **service account key** (`service_account.json`) grants full admin access. Keep it on your machine only.
- License keys are stored as SHA-256 hashes in Firestore. Even if someone accesses the database, they can't recover the original keys.
- Machine binding is enforced both in Firestore rules (one-time write) and in the app's local validation logic.

## Firestore Document Structure

Each license is stored at `licenses/{sha256_hash_of_key}`:

| Field        | Type      | Description                                    |
|-------------|-----------|------------------------------------------------|
| key_preview | string    | Last 4 characters of the key (for admin display) |
| machine_id  | string    | SHA-256 of the bound machine's ID (empty = unbound) |
| expires_at  | timestamp | When the license expires                        |
| revoked     | boolean   | Whether the license has been revoked             |
| created_at  | timestamp | When the license was generated                   |
| note        | string    | Admin notes (customer name, etc.)                |
