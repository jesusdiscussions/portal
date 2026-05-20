# portal

Community comment emailer (Blog Platform Engine).

## GitHub Pages files

Deploy these to your Pages site root (same folder as `posts/`):

- `index.html` — confirm / unsubscribe
- `CommentsPage.html` — comments portal
- `go.html` — **mail handoff** (Yahoo / in-app browsers). If someone opens a long `go.html?b=…` link from the portal “stuck in app” banner, this page helps open Safari/Chrome with the real URL.

**Email links (confirm / unsubscribe / join comments)** are wrapped as a **short** Apps Script URL `…/exec?go=TOKEN`. The real GitHub Pages URL is stored on the spreadsheet tab `_MailHandoff` (created by **Setup sheets**) so Yahoo Mail is less likely to **truncate** the link (truncation used to show “bad link” or loop inside the mail app). After `clasp push`, **redeploy** the web app and run **Setup sheets** once if `_MailHandoff` does not exist yet.

Without `go.html`, long fallback handoff links from the portal can 404 if not deployed.
