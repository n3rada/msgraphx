# 🔭 msgraphX

Microsoft Graph eXploitation toolkit. ~~Ab~~using the [Microsoft Graph API](https://learn.microsoft.com/en-us/graph/overview) [SDK](https://github.com/microsoftgraph/msgraph-sdk-python) to search and harvest SharePoint files, Outlook mail, Teams messages and Microsoft 365 data during red team operations and penetration tests.

- **SharePoint**: search across all sites or a specific site (`--site`), filter by filetype, predefined hunt queries, bulk download, enumerate sites via group membership
- **Outlook**: build communication graphs from mailboxes, KQL keyword search, download emails as `.eml`
- **Teams**: search DMs and channel messages, browse chats, send messages to any user or channel
- **Entra ID**: enumerate users, groups, devices, app registrations (with credential/secret expiry status), service principals, roles, PIM assignments, OAuth2 grants, conditional access policies, MFA methods
- **Groups**: list M365 Unified groups, enumerate members (direct or transitive), resolve group SharePoint sites
- **Me**: OneDrive tree view, upload, download, calendar, OneNote, Planner tasks, people graph, trending/shared/used documents
- **Output**: local caching of search results per identity, JSON/NDJSON export, resumable downloads

For the full command reference see [USAGE.md](USAGE.md).

## 📦 Installation

```bash
uv tool install git+https://github.com/n3rada/msgraphx.git
```

Run without installing:

```bash
uvx --from git+https://github.com/n3rada/msgraphx.git msgraphx --help
```

With pipx or pip:

```bash
pipx install 'git+https://github.com/n3rada/msgraphx.git'
pip install 'git+https://github.com/n3rada/msgraphx.git'
```

## 🔑 Authentication

### Delegated (user context)

Use [msauth-browser](https://github.com/n3rada/msauth-browser) to authenticate through a real browser and save tokens:

```bash
msauth-browser --save roadtools
msgraphx outlook contacts   # picks up .roadtools_auth automatically
```

Or pass the token directly:

```bash
msgraphx --access-token <JWT> sp search "password"
export ACCESS_TOKEN=<JWT>
msgraphx sp search "password"
```

### Application (app-only)

```bash
msgraphx --tenant-id <tid> --client-id <cid> --client-secret <secret> sp search "password"
```

Or via environment variables `TENANT_ID`, `CLIENT_ID`, `CLIENT_SECRET`.

> [!NOTE]
> App-only tokens cannot use the `/me` endpoint. Modules that require a user context (Outlook, me) only work with delegated auth.

## ⚔️ Quick start

```bash
# SharePoint: hunt for credentials across all sites
msgraphx sp search --hunt credentials

# Outlook: build a communication graph
msgraphx outlook contacts

# Teams: search DMs for AWS keys
msgraphx teams chat "aws key" --after 90d

# Entra ID: enumerate admin accounts
msgraphx aad search --hunt admins

# Entra ID: list PIM assignments
msgraphx aad pim

# Entra ID: find apps with active secrets or never-expiring credentials
msgraphx aad apps --with-secrets
msgraphx aad app <client-id-or-object-id>

# Me: dump OneNote pages
msgraphx me onenote --list

# Me: browse OneDrive tree
msgraphx me drive tree

# Groups: list your M365 groups
msgraphx groups list --mine

# Teams: send a message to a user
msgraphx teams send --to target@corp.com "message"
```

See [USAGE.md](USAGE.md) for all flags, subcommands, and options for each module.

## 🔬 Graph Explorer

[Graph Explorer](https://developer.microsoft.com/en-us/graph/graph-explorer) lets you run live queries against the Graph API and generate ready-to-paste Python SDK code snippets via the *Code snippets* tab. Use it to prototype any query before implementing a module.
