Microsoft Graph eXploitation toolkit.

## Installation

Install with `pipx` to keep the environment isolated:
```shell
pipx install "msgraphx@git+https://github.com/n3rada/msgraphx.git"
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