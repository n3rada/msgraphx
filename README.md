# 🔭 msgraphX

Microsoft Graph eXploitation toolkit. ~~Ab~~using the [Microsoft Graph API](https://learn.microsoft.com/en-us/graph/overview) SDK to search and harvest SharePoint files, Outlook mail, Teams messages and Microsoft 365 data during red team operations and penetration tests.

- **SharePoint**: search across all sites, filter by filetype, use predefined hunt queries, bulk download
- **Outlook**: build communication graphs from mailboxes, KQL keyword search, download emails as `.eml`
- **Teams**: search DMs and channel messages, browse chats, inspect message context
- **Authentication**: delegated (user token via OAuth PKCE / [msauth-browser](https://github.com/n3rada/msauth-browser)) and app-only (client credentials / service principal)
- **Output**: local caching of search results, JSON export, resumable downloads

## 📦 Installation

Prefer using [`uv`](https://docs.astral.sh/uv/), a fast Python package manager that installs tools in isolated environments. Alternatively, [`pipx`](https://pypa.github.io/pipx/) or `pip` work as well.

### With [uv](https://docs.astral.sh/uv/)

[`uv tool install`](https://docs.astral.sh/uv/guides/tools/#installing-tools) persistently installs the tool and adds it to your `PATH`, similar to `pipx`:

**From GitHub (latest):**

```bash
uv tool install git+https://github.com/n3rada/msgraphx.git
```

After installation, `msgraphx` is available directly:

```bash
msgraphx --help
```

To upgrade later:

```bash
uv tool upgrade msgraphx
```

> [!TIP]
> You can also run `msgraphx` **without installing** it using [`uvx`](https://docs.astral.sh/uv/guides/tools/#running-tools), which creates a temporary isolated environment on the fly:
> ```bash
> uvx --from git+https://github.com/n3rada/msgraphx.git msgraphx --help
> ```

### With pipx or pip

```bash
pipx install 'git+https://github.com/n3rada/msgraphx.git'
```

```bash
pip install 'git+https://github.com/n3rada/msgraphx.git'
```

## 🔑 Authentication

`msgraphx` supports two authentication modes.

### Delegated (user context)

Required for any module that acts on behalf of a user (e.g., Outlook, SharePoint search as yourself).

The easiest way to get a valid token is [msauth-browser](https://github.com/n3rada/msauth-browser), which drives a real Chromium browser through the full OAuth PKCE flow and handles MFA, Conditional Access, and CAPTCHAs transparently:

```bash
# authenticate as Graph Explorer and save tokens in .roadtools_auth
msauth-browser --save roadtools
```

Then run `msgraphx` from the same directory - it will pick up `.roadtools_auth` automatically:

```bash
msgraphx outlook contacts
```

Alternatively, pass the token directly or via environment variable:

```bash
# via flag
msgraphx --access-token <JWT> outlook contacts

# via env var
export ACCESS_TOKEN=<JWT>
msgraphx outlook contacts
```

### Application (app-only)

For app-only flows (service principals with client credentials):

```bash
msgraphx --tenant-id <tid> --client-id <cid> --client-secret <secret> sp search "password"
```

Or via environment variables:

```bash
export TENANT_ID=<tid>
export CLIENT_ID=<cid>
export CLIENT_SECRET=<secret>
msgraphx sp search "password"
```

> [!NOTE]
> App-only tokens cannot use the `/me` endpoint. Modules that require a user context (Outlook) only work with delegated auth.


## ⚔️ Modules

> [!NOTE]
> All modules default to the **last year** (`--after 1y`). Pass `--all` to remove the time bound, or `--after`/`--before` to set a custom range. These are global flags available to every subcommand.

### 📧 Outlook

#### Contacts (connection graph)

Build a full communication graph from your mailbox. By default analyses both sent and received mail across four ranked tables:

- 📤 **Sent → To**: who you direct-email most
- 📤 **Sent → CC**: who you copy most
- 📥 **Received → as To**: who emails you directly most
- 📥 **Received → as CC**: who copies you most

```shell
msgraphx outlook contacts
# or
msgraphx mail contacts
```

Restrict to sent or received only:
```shell
msgraphx mail contacts --only sent
msgraphx mail contacts --only received
```

Fetch everything with no time bound:
```shell
msgraphx mail contacts --all
```

Limit to the last 90 days and show top 50:
```shell
msgraphx mail contacts --after 90d --top 50
```

Save the full ranked list as JSON:
```shell
msgraphx mail contacts --save /tmp/contacts.json
```

#### Search

Search your mailbox using KQL. Streams up to 1 000 results (Exchange API cap):

```shell
msgraphx outlook search "password"
msgraphx outlook search --from alice@corp.com
msgraphx outlook search --subject "VPN" --has-attachments
msgraphx outlook search "credentials" --after 90d
```

Results are cached locally (`~/.local/share/msgraphx/last_mail.json`).

#### Download

Download specific emails as `.eml` files from the last search:

```shell
# First, search
msgraphx outlook search "password"
#    1.  RE: VPN config  alice@corp.com  2025-03-12
#    2.  FW: passwords   bob@corp.com    2025-01-08

# Then, download by index
msgraphx outlook download 1
msgraphx outlook download 1,2
msgraphx outlook download 1-5 --save /tmp/loot/
```

`.eml` files open natively in most mail clients.

### 🏢 SharePoint

#### Search

Search for anything across SharePoint. Defaults to the last year:

```shell
msgraphx sp search "password"
```

Filter by filetype:
```shell
msgraphx sp search --filetype pdf
msgraphx sp search -f docx "confidential"
```

Use predefined hunt queries:
```shell
msgraphx sp search --hunt credentials
msgraphx sp search --hunt ssh --all
msgraphx sp search --hunt office --after 90d
```

Search only within your own Microsoft 365 groups:
```shell
msgraphx sp search "password" --my-groups
```

Save all results to disk:
```shell
msgraphx sp search --hunt credentials --save /tmp/loot/
```

#### Download

Download specific files from the last search by index:

```shell
# First, search
msgraphx sp search "Itron" --filetype pdf
#    1.  Itron_Report_2024.pdf  jsmith  2.1 MB  2024-11-03
#    2.  Itron_Specs.pdf        jdoe    840 KB  2024-09-15
#    3.  Itron_Invoice.pdf      admin   120 KB  2025-01-20

# Then, download by index
msgraphx sp download 2
msgraphx sp download 1,3
msgraphx sp download 1-3
```

Each search caches results locally (`~/.local/share/msgraphx/last_sharepoint.json`). A new search always replaces the previous cache.

Download to a specific directory:
```shell
msgraphx sp download 1-3 --save /tmp/loot/
```

Full drive dump (requires `--drive-id`):

```shell
msgraphx --drive-id <drive-id> sp download --save /tmp/loot/
```

Resume interrupted downloads (default behaviour, skips files already present with matching size):
```shell
msgraphx --drive-id <drive-id> sp download --save /tmp/loot/
```

Force re-download everything:
```shell
msgraphx --drive-id <drive-id> sp download --no-resume --save /tmp/loot/
```


### 💬 Teams

Requires delegated auth. Both subcommands use `POST /search/query` (`EntityType.ChatMessage`) and need **`ChannelMessage.Read.All`** (admin consent) in addition to `Chat.Read`.

#### Chat (personal messages)

Search 1:1 DMs and group chats:

```shell
# keyword search
msgraphx teams chat "password"
msgraphx teams chat "vpn credentials"

# sender filter (client-side)
msgraphx teams chat "budget" --from alice

# date range
msgraphx teams chat "aws key" --after 90d
msgraphx teams chat "deploy" --after 2024-01-01 --before 2024-06-01

# wildcard: return everything
msgraphx teams chat
```

Results are cached locally (`~/.local/share/msgraphx/last_teams.json`).

#### Channel (workspace channels)

Search messages across all Teams channels you have access to:

```shell
msgraphx teams channel "password"
msgraphx teams channel "from:alice@corp.com"
msgraphx teams channel "incident" --after 30d
```

KQL is passed directly to the Search API, so any valid KQL expression works:

```shell
msgraphx teams channel "subject:deployment AND azure"
```

#### Show

Show context around a cached result, or browse a named chat directly. The argument
auto-detects the mode: a number/range opens the cache, any other string finds a chat.

```shell
# cached result (index from last search)
msgraphx teams show 3
msgraphx teams show 1-5
msgraphx teams show 2 --context 8

# browse a named chat (topic or member name match, last 20 by default)
msgraphx teams show alice
msgraphx teams show "project phoenix" --last 50
msgraphx teams show alice --last 5
```

## 🔬 Graph Explorer

[Graph Explorer](https://developer.microsoft.com/en-us/graph/graph-explorer) is Microsoft's interactive API sandbox. It lets you run live queries against the Graph API, inspect raw responses, and generate ready-to-paste **Python SDK code snippets** for any request via the *Code snippets* tab.

Use it to prototype a query before implementing it as a module, or to understand exactly what fields a response contains.

