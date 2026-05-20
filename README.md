# portal

Community comment emailer (Blog Platform Engine).

## GitHub Pages files

Deploy these to your Pages site root (same folder as `posts/`):

- `index.html` — confirm / unsubscribe
- `CommentsPage.html` — comments portal
- `go.html` — **mail handoff** (Yahoo / in-app browsers). Email links are wrapped to hit this page first, then open Safari/Chrome with the real URL.

Without `go.html`, wrapped links from email will 404.
