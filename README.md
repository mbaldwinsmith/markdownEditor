# Markdown Editor PWA

A zero-build Progressive Web App Markdown editor designed to run directly from GitHub Pages. It offers streamlined editing, formatting shortcuts, offline support, and Google Drive integration for opening and saving `.md` files.

## Features

- **Streamlined editor** – focus on Markdown syntax with helpful toolbar shortcuts.
- **Formatting toolbar** – buttons for bold, italic, headings, lists, links, images, tables, and horizontal rules.
- **Word & character counts** – update live as you type.
- **Offline-ready** – installable PWA with caching via a service worker.
- **Google Drive sync** – sign in with your Google account to open and save Markdown files.

## Getting started

1. Clone or fork this repository.
2. Enable GitHub Pages for the repository (Settings → Pages → Deploy from branch → `main`).
3. Visit the published site – everything runs client-side, no build pipeline required.

To work locally without a server, simply open `index.html` in a browser. For full PWA behaviour you should use a local HTTP server (for example `python -m http.server`) so that service workers can register.

## Google Drive configuration

Google requires that you supply your own OAuth client ID (and optionally an API key) for Drive access. Configure these values via the meta tags in `index.html`.

1. Visit the [Google Cloud Console](https://console.cloud.google.com/).
2. Create a project (or choose an existing one) and enable the **Google Drive API**.
3. Create credentials:
   - **OAuth 2.0 Client ID** (type: Web application). Add your GitHub Pages origin to the authorised JavaScript origins.
   - **API key** *(optional)*. If you choose to use one, restrict it to the origin that will host the editor and copy it into the `google-api-key` meta tag.
4. Update the placeholder value in the `<meta name="google-oauth-client-id">` tag inside `index.html` with your client ID. If you created an API key, place it in the `<meta name="google-api-key">` tag.
5. Deploy the updated files. When you open the editor, click **Sign in** to authorise Google Drive access, then use **Open**, **Save**, or **Save as** to work with Drive files.

> ⚠️ The app requests the `drive.file` and `drive.readonly` scopes. Ensure your OAuth consent screen is configured for external users if you plan to share the app.

## Project structure

```
.
├── index.html          # Main application markup
├── styles.css          # Layout and visual styling
├── app.js              # Editor logic and Google Drive integration
├── manifest.json       # PWA manifest
├── service-worker.js   # Offline caching
└── icons/              # PWA icons
```

## Development notes

- The app persists the current document and last opened file in the browser's `localStorage`. Google OAuth configuration is provided via static meta tags.
- Offline support caches the static assets; Google Drive actions require an active internet connection.
- Because everything is static, you can fork and customise the UI without setting up a toolchain.

## License

[MIT](LICENSE)
