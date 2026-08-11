# Mark's Markdown Editor

A zero-build Progressive Web App Markdown editor designed to run directly from GitHub Pages. It offers streamlined editing, formatting shortcuts, offline support, and Google Drive integration for opening and saving `.md` files.

## Features

- **Streamlined editor** – focus on Markdown syntax with helpful toolbar shortcuts.
- **Formatting toolbar** – buttons for bold, italic, headings, lists, links, images, tables, and horizontal rules.
- **Word & character counts** – update live as you type.
- **Offline-ready** – installable PWA with caching via a service worker.
- **Google Drive sync** – sign in with your Google account to open and save Markdown files.

## Technologies used

- **Vanilla JavaScript** – implements the editor logic, toolbar behaviour, and Google Drive integration without a build step.
- **HTML5** – provides the static structure for the editor and supporting pages such as privacy and terms documents.
- **CSS3** – styles the interface to deliver a clean, responsive editing experience across devices.
- **Progressive Web App APIs** – manifest and service worker files enable offline usage and installation prompts.
- **Google Drive API** – powers authentication and remote file operations when users connect their Google accounts.

## Getting started

1. Clone or fork this repository.
2. Enable GitHub Pages for the repository (Settings → Pages → Deploy from branch → `main`).
3. Visit the published site – everything runs client-side, no build pipeline required.

To work locally without a server, simply open `index.html` in a browser. For full PWA behaviour you should use a local HTTP server (for example `python -m http.server`) so that service workers can register.

## Google Drive configuration

Google Drive sign-in requires a browser OAuth client ID. Opening arbitrary existing files through Google Picker also requires a browser-restricted API key and the Google Cloud project number (the Picker app ID). The app attempts to load `./config/google-drive.json`, which should return `clientId`, `apiKey`, and `appId` fields. Browser client IDs, restricted API keys, and project numbers are public identifiers in a static web app; never add an OAuth client secret.

1. Visit the [Google Cloud Console](https://console.cloud.google.com/).
2. Create a project (or choose an existing one) and enable the **Google Drive API**.
3. Create credentials:
   - **OAuth 2.0 Client ID** (type: Web application). Add your GitHub Pages origin to the authorised JavaScript origins.
   - **API key**. Restrict it to the Google Picker API and to the exact origins that host the editor.
   - **Project number**. Find it under **IAM & Admin → Settings**; Google Picker calls this the app ID.
4. Configure your hosting platform to expose the credentials at `/config/google-drive.json` or an equivalent authenticated endpoint. The payload should resemble:

   ```json
   {
     "clientId": "YOUR_CLIENT_ID",
     "apiKey": "YOUR_BROWSER_RESTRICTED_API_KEY",
     "appId": "YOUR_GOOGLE_CLOUD_PROJECT_NUMBER"
   }
   ```

   Do not place an OAuth client secret in this file or anywhere in the browser app.
5. For local development or static hosting without a config endpoint, you may provide `google-oauth-client-id`, `google-api-key`, and `google-cloud-project-number` meta tags in `index.html`. The project number can also be derived from the standard browser client ID format.
6. Deploy the updated files and configuration. Click **Sign in**, then choose **Open → My Drive** to select an existing Markdown file with Google Picker. **Save** and **Save as** continue to create or update files authorised for this editor.

> The app requests only `drive.file`, a non-sensitive per-file scope. Google Picker grants the editor access only to files you explicitly select, while the editor can continue to update files it created.

## Project structure

```
.
├── index.html          # Main application markup
├── privacy.html        # Privacy Policy for the web app
├── styles.css          # Layout and visual styling
├── app.js              # Editor logic and Google Drive integration
├── manifest.json       # PWA manifest
├── service-worker.js   # Offline caching
├── terms.html          # Terms of Service for the web app
└── icons/              # PWA icons
```

## Policies

- [Privacy Policy](privacy.html)
- [Terms of Service](terms.html)

## Development notes

- The app persists the current document and last opened file in the browser's `localStorage`. Google OAuth configuration is loaded from a secure runtime endpoint when available (falling back to the meta tag for local development).
- Offline support caches the static assets; Google Drive actions require an active internet connection.
- Because everything is static, you can fork and customise the UI without setting up a toolchain.

## License

[MIT](LICENSE)
