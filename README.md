Microsoft Graph eXploitation toolkit.

## Installation

Install with `pipx` to keep the environment isolated:
```shell
pipx install graphx
```

## 🔐 Authentication

GraphX provides a Playwright-based authenticator to acquire a Microsoft Graph access token.

```shell
graphx auth
```

Options:
- `--prt-cookie <JWT>`: Use an `x-ms-RefreshTokenCredential` PRT cookie for SSO-based login.
- `--headless`: Run Playwright in headless mode.

This print the access token on the `stdout` and generate a `.roadtools_auth` file containing the access and refresh tokens.

You can choose to install directly the standalone `graphx-auth`:
```shell
pipx install 'graphx[auth]'
```

## ⚔️ Exploitation Modules

GraphX is built with modularity in mind. Each module is dynamically loaded and can be invoked as a subcommand.


## Graph Explorer 101

Microsoft's Graph Explorer (`de8bc8b5-d9f9-48b1-a8ad-b748da725064`) is a public client (no secret) and allows OAuth2-based interaction with Graph. Thus, it is not FOCI-enabled. 

### FOCI (Family of Client IDs)

FOCI allows Microsoft apps to share refresh tokens. Graph Explorer is not part of such a family, meaning its refresh token is not reusable across other clients.


### About the PRT Cookie

The so-called PRT cookie is officially `x-ms-RefreshTokenCredential` and it is a JSON Web Token (JWT). The actual Primary Refresh Token (PRT) is encapsulated within the `refresh_token`, which is encrypted by a key under the control of Entra ID, rendering its contents opaque. 

It can be used as a cookie wired to `login.microsoftonline.com` domain in order to use-it to authenticate to the service while skiping credential prompts.