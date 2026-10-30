# Repository instructions

## Project

msgraphX is a Python 3.12 Microsoft Graph post-exploitation toolkit. It uses the official Microsoft Graph Python SDK to search and enumerate Microsoft 365 services.

Read [`DEVELOPMENT.md`](DEVELOPMENT.md) when changing architecture, Graph requests, module behavior, authentication, pagination, caching, error handling, logging, output, or dependencies. For focused changes, read only the relevant section and its linked source files.

Before preparing a commit, pull the current branch with `git pull --rebase --autostash`; never commit against stale remote history. Then follow the human contribution and commit conventions in [`DEVELOPMENT.md`](DEVELOPMENT.md).

## Core contracts

- Use the official Graph SDK for `graph.microsoft.com`; raw `httpx` is limited to the documented non-Graph endpoints.
- Pass the shared `GraphContext`; do not construct module-local Graph clients.
- Keep module entry points, Graph error handling, pagination utilities, and cache-to-index-to-fetch behavior consistent.
- Keep JSON output as the only stdout content under `--json`; use Loguru for stderr progress and the shared Rich console for human output.
- Target Python 3.12, preserve async SDK behavior, and use concrete modern type annotations.
- `verify=False` is intentional for intercepting proxies and self-signed assessment environments.

## Verification

Run affected tests, or the full suite with `uv run pytest -v`.
