# Blog Platform Engine — Phase 1 (local)

Phase 1 adds the V2 spreadsheet schema, User ID assignment, and PostLog migration. It does **not** include Worker, V2 email send, merge tags, or frontend threading yet.

## New menu items

| Menu item | Purpose |
|-----------|---------|
| **Setup V2 sheets (one-shot)** | Creates `PageContent`, `EmailQueue`, `ActionTokens`, ensures `CommentsLog` headers; never touches **Notes** |
| **Migrate PostLog → V2…** | Master rows → `PageContent`, subscribers → `EmailQueue`, assigns User IDs, **deletes** `PostLog` / `CampaignLog` |
| **Assign missing User IDs** | Fills blank User ID cells in `EmailQueue` from the next sequence number |

Legacy **Run Post Send & Sync** remains but alerts if `PostLog` is gone.

## Tabs

### PageContent
`Post ID`, `Live`, `Page Headline`, `Page Body HTML`, `Google Drive File IDs`, `Comments Duration Open`, `Discussion Start At`, `Force Closed`, `Last Updated`, `Publish Version`

### EmailQueue
`User ID`, `Post ID`, `Email Address`, `First Name`, `Last Name`, `STATUS`, `Email Subject`, `Email Body`, `Send?`, `Sent Timestamp`

### ActionTokens
`Token`, `Post ID`, `Email Address`, `Action`, `Expires At`, `Used` (empty until Phase 2)

### CommentsLog
Unchanged structure; **not cleared** on migration.

### Notes
**Whitelisted** — never deleted or overwritten by setup/migration.

## User IDs

- Numeric strings starting at **1000** (`1000`, `1001`, …)
- **One unique User ID per EmailQueue row** (same email on different Post IDs gets different IDs)
- Next ID stored in Script Property `NEXT_USER_ID`

## Migration behavior

- Reads **PostLog** or **CampaignLog**
- Master row (`First Name` = `post id`) → **PageContent** (uses Launch/Modify routing: `Last Name` = web target when different)
- Subscriber rows → **EmailQueue** with master’s subject/body as initial email copy
- **Deletes** legacy PostLog tab after success
- Skips clusters without a valid master row (reported in summary alert)

## Deploy

```bash
clasp push
```

Then in the spreadsheet: **Blog Platform Engine → Setup V2 sheets**, then **Migrate PostLog → V2** if you have existing data.
