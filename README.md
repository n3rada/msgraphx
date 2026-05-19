# 🔭 msgraphx

Microsoft Graph eXploitation toolkit.

## 📦 Installation

Prefer using [`uv`](https://docs.astral.sh/uv/), a fast Python package manager that installs tools in isolated environments. Alternatively, [`pipx`](https://pypa.github.io/pipx/) or `pip` work as well.

### With [uv](https://docs.astral.sh/uv/) (recommended)

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

## ⚔️ Modules

### SharePoint

#### Search

You can search for anything such as `password`:
```shell
msgraphx sp search "password"
```

You can specify the filetype using `--filetype` or `-f`:
```shell
msgraphx sp search --filetype pdf
msgraphx sp search -f docx "confidential"
```

You can use predefined hunt queries for credentials from last 365 days (1 year):
```shell
msgraphx sp search --hunt credentials --after 365d --save "/tmp/"
```