# BM Audit Admin Launcher (GitHub Pages)

## Before publishing
1. Open `index.html`.
2. Find `const ADMIN_APP_URL = "PASTE_YOUR_BM_AUDIT_ADMIN_URL_HERE";`
3. Replace the placeholder with your real BM Audit Admin app URL (HTTPS).
4. Replace `icons/bm-audit-192.png` and `icons/bm-audit-512.png` with your approved BM Audit logo exported as square PNGs (192x192 and 512x512).

## Publish on GitHub Pages
1. Upload `index.html`, `manifest.json`, `service-worker.js`, and the `icons` folder to the root of your BM Audit GitHub Pages repository (or to the matching Pages folder).
2. In repository Settings → Pages, enable the intended branch/folder.
3. Open the HTTPS GitHub Pages URL.
4. Test on mobile Chrome. Use the browser's Install app option or the install button when available.
5. On iPhone/iPad, open in Safari → Share → Add to Home Screen.

## Important security note
This is a static launcher, not an authentication system. Hiding a PIN field does not protect the backend. The actual admin app / Apps Script deployment must enforce trusted-user access server-side. Do not put PINs, passwords, API keys, or private spreadsheet credentials in this repository.

## Install button behavior
The install button appears when the browser offers the PWA installation prompt. It hides in installed standalone mode and after `appinstalled`. On iOS, browsers do not expose the same install prompt; the button displays Add to Home Screen instructions.
