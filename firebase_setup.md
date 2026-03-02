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

These rules allow:
- **Admin (Adnan@thothica.com)**: Full read/write/delete via Firebase Auth
- **App clients**: Read license docs (key hash is the secret) + one-time machine binding
- **Everyone else**: Denied

## 4. Enable Firebase Authentication

1. Go to **Build → Authentication**
2. Click **Get started**
3. Go to the **Sign-in method** tab
4. Click **Google** → toggle **Enable** → set a support email → click **Save**
5. Go to **Settings → Authorized domains** and verify `localhost` is listed (it should be by default)

## 5. Get Your API Key and Config

1. Go to **Project Settings** (gear icon) → **General**
2. Scroll down to **Your apps**
3. Click **Add app** → choose **Web** (</> icon)
4. Name it (e.g., `pdf-toolkit-client`)
5. From the config snippet, note these values:
   - `apiKey` (e.g., `AIzaSy...`)
   - `authDomain` (e.g., `pdf-toolkit-abc123.firebaseapp.com`)
   - `projectId` (e.g., `pdf-toolkit-abc123`)

## 6. Update Your Config Files

### `.env` (for the desktop app — `license.py` reads this):
```
FIREBASE_PROJECT_ID=pdf-toolkit-abc123
FIREBASE_API_KEY=AIzaSy...
FIREBASE_AUTH_DOMAIN=pdf-toolkit-abc123.firebaseapp.com
ADMIN_EMAIL=Adnan@thothica.com
```

### `admin/index.html` (for the admin portal):
Open `admin/index.html` and replace the placeholder values near the top of the `<script>` section:
```javascript
const FIREBASE_CONFIG = {
  apiKey:      "AIzaSy...",
  authDomain:  "pdf-toolkit-abc123.firebaseapp.com",
  projectId:   "pdf-toolkit-abc123",
};
```

## 7. (Optional) Create a Service Account for CLI

The admin portal (`admin/index.html`) is now the primary way to manage licenses.
The CLI (`admin_license.py`) is optional, for scripting/automation.

If you want to use the CLI:
1. Go to **Project Settings → Service Accounts**
2. Click **Generate new private key**
3. Save the downloaded JSON file as `service_account.json` in the project directory
4. **IMPORTANT:** Never include this file in the EXE or share it with customers
5. Install: `pip install firebase-admin`

## 8. Test the Setup

### Open the admin portal:
1. Start a local server: `python -m http.server 8080 --directory admin`
2. Open `http://localhost:8080` in your browser
3. Sign in with your Google account (Adnan@thothica.com)
4. Click **+ Generate Key** to create a test license

### Test in the desktop app:
1. Run `python app.py`
2. The license modal should appear
3. Enter the key you generated in the admin portal
4. The app should activate and load normally

### Test revocation:
1. In the admin portal, click **Revoke** next to the test license
2. Restart the app — it should show "This license has been revoked" on the next online check

### Test unbinding:
1. In the admin portal, click **Unbind** next to a bound license
2. The key can now be activated on a different machine

## 9. (Optional) Deploy Admin Portal to Firebase Hosting

To access the admin portal from anywhere (not just localhost):

1. Install Firebase CLI: `npm install -g firebase-tools`
2. Login: `firebase login`
3. Initialize hosting: `firebase init hosting`
   - Select your project
   - Set public directory to `admin`
   - Configure as single-page app: **No**
4. Deploy: `firebase deploy --only hosting`
5. Your admin portal will be live at `https://your-project-id.web.app`

Make sure to add your hosting domain to **Authentication → Settings → Authorized domains**.

## Security Notes

- The **API key** is safe to embed in both the app and admin portal. It only identifies the Firebase project; all access control is enforced by Firestore security rules.
- The **service account key** (`service_account.json`) grants full admin access. Keep it on your machine only. Not needed if you only use the admin portal.
- **Admin access** is restricted to `Adnan@thothica.com` via Firestore security rules. No other Google account can create, modify, or delete licenses.
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
