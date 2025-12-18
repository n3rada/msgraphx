Microsoft Graph eXploitation toolkit.

## Installation

Install with `pipx` to keep the environment isolated:
```shell
pipx install "msgraphx@git+https://github.com/n3rada/msgraphx.git"
```

## ⚔️ Exploitation Modules

msgraphx is built with modularity in mind. Each module is dynamically loaded and can be invoked as a subcommand.


### SharePoint

#### Search

```shell
msgraphx sp search "nuc" - search for "nuc"
msgraphx sp search - defaults to "*" (search all)
msgraphx sp search filetype:pdf - search for PDFs
msgraphx sp search "nuc" --filetype pdf - combine positional query with filetype filter
```