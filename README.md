# 🔭 msgraphx

Microsoft Graph eXploitation toolkit.

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

- 📤 **Sent → To** — who you direct-email most
- 📤 **Sent → CC** — who you copy most
- 📥 **Received → as To** — who emails you directly most
- 📥 **Received → as CC** — who copies you most

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

Each search caches results locally (`~/.local/share/msgraphx/last_search.json`). A new search always replaces the previous cache.

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