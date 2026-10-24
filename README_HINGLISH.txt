BATTERY AUDIT TRACKER — GITHUB REPOSITORY LAYOUT
=================================================

This folder structure deliberately matches the existing RM Audit App style:
- index.html                     central launcher
- admin/                         Admin UI + config + PWA manifest/service worker
- technician/                    Technician UI + config + PWA manifest/service worker
- assets/                        separate Admin and Technician SVG logos
- apps-script-backend/           source reference only; NOT deployed automatically

Upload the CONTENTS of Battery_Audit_Tracker_GitHub_Ready into the ROOT of the
battery-audit-tracker-admin repository only if you want one central repo/launcher.
If using two separate GitHub repositories as planned:
1) Admin repository root: upload contents of admin/ plus assets/battery-admin-logo.svg
   (and adjust logo paths to root if you move files; easiest is use the full combined
   repository as a single Pages site first, then split later.)
2) Technician repository root: upload contents of technician/ plus assets/battery-technician-logo.svg.

IMPORTANT: The Admin and Technician pages are wrappers around the Apps Script Web App.
Before publishing, confirm their iframe/login behaviour in a browser. The Apps Script
URL is configured in config.js. Token access must be provided in the technician URL.
Do not put the Admin PIN in GitHub. Do not run backend setup functions without backup.
No live Google Sheet was accessed or changed while preparing this package.
