# msgraphX: AI Guidance

This file is the canonical AI guidance for this repository, shared across Claude, Copilot, and other agents.

## Context

msgraphX is a red team / penetration testing tool. It exploits the Microsoft Graph API to harvest SharePoint files, Outlook mail, Teams messages, and Microsoft 365 data during offensive operations. It is in active beta — breaking changes are expected and backward compatibility is never a concern.

`verify=False` on HTTP calls is intentional. Operators work against environments with intercepting proxies and self-signed certificates. Never flag it.

## Read Order

1. [README.md](README.md) — CLI behavior, authentication modes, all module usage.
2. [DEVELOPMENT.md](DEVELOPMENT.md) — architecture, design principles, SDK rules, and extension model (source of truth).
3. [src/msgraphx/cli.py](src/msgraphx/cli.py) — runtime flow, auth bootstrap, argument dispatch.
4. [tests/](tests/) — expected behavior and regression boundaries.

## Commands

```bash
# Install in editable mode
uv sync

# Run all tests
uv run pytest -v

# Run a single test file
uv run pytest tests/test_cache.py -v

# Run a single test by name
uv run pytest tests/test_cache.py::test_parse_indices -v

# Run the CLI (development)
uv run msgraphx --help
uv run msgraphx --trace outlook contacts
```

## Source Map

| Path | Purpose |
|------|---------|
| [`src/msgraphx/cli.py`](src/msgraphx/cli.py) | Entry point, argument dispatch, auth bootstrap |
| [`src/msgraphx/core/context.py`](src/msgraphx/core/context.py) | `GraphContext` — shared runtime state |
| [`src/msgraphx/core/graph_search.py`](src/msgraphx/core/graph_search.py) | `POST /search/query` wrapper (`SearchOptions`, `search_entities`) |
| [`src/msgraphx/modules/`](src/msgraphx/modules/) | Per-service capability packages |
| [`src/msgraphx/utils/pagination.py`](src/msgraphx/utils/pagination.py) | `GraphPaginator`, `collect_all` — all paginated SDK calls go here |
| [`src/msgraphx/utils/errors.py`](src/msgraphx/utils/errors.py) | `handle_graph_errors`, `AuthenticationError`, `ForbiddenGraphError` |
| [`src/msgraphx/utils/cache.py`](src/msgraphx/utils/cache.py) | XDG result cache, `parse_indices` |
| [`src/msgraphx/utils/tokens.py`](src/msgraphx/utils/tokens.py) | Token loading, refresh, `.roadtools_auth` interop |
| [`src/msgraphx/utils/logbook.py`](src/msgraphx/utils/logbook.py) | Loguru setup, log rotation |
| [`src/msgraphx/utils/console.py`](src/msgraphx/utils/console.py) | Shared Rich `Console` instance (stdout) |
| [`src/msgraphx/utils/output.py`](src/msgraphx/utils/output.py) | `print_json(data)` — writes JSON to stdout for `--json` mode |
| [`tests/`](tests/) | pytest suite |

## Architecture and Design Principles

This section is a strict summary. Detailed rationale and examples live in [DEVELOPMENT.md](DEVELOPMENT.md).
If this file and [DEVELOPMENT.md](DEVELOPMENT.md) ever diverge, follow [DEVELOPMENT.md](DEVELOPMENT.md).

Enforce these rules on every change:

**1. GraphContext is the only shared state.** Modules receive a `GraphContext` and must not construct their own `GraphServiceClient`. Authentication, scopes, and cached user identity all come from `context`.

**2. Module contract is fixed.** Every module and sub-module exposes exactly `add_arguments(parser, parents)` and `async run_with_arguments(context, args) -> int`, decorated with `@handle_graph_errors`. Do not invent other entry points.

**3. Two distinct API paths — use the right one.**
- `core/graph_search.py` (`search_entities`) for full-text search: SharePoint (`DriveItem`), Teams (`ChatMessage`), Outlook (`Message`). Handles pagination and the Exchange 1 000-hit cap internally.
- SDK list builders (`client.users`, `client.groups`, `client.drives`, etc.) for enumeration, always via `pagination.GraphPaginator` or `pagination.collect_all`. Never hand-roll `@odata.nextLink` loops.

**3a. Never use raw httpx for Graph API calls.**
The SDK must be used for every `graph.microsoft.com` call. httpx is only permitted for non-Graph endpoints the SDK does not cover: Azure Blob upload session chunk PUT requests (`me/drive.py`), `login.microsoftonline.com` OAuth token refresh (`utils/tokens.py`), and `api.ipify.org` public IP check (`cli.py`). For parallel Graph calls use `asyncio.gather`; for batch use `context.graph_client.batch()`. See [DEVELOPMENT.md](DEVELOPMENT.md) for examples.

**4. Cache → index → fetch is the download pattern.** Search modules write `cache.save_results(items, key)`. Download modules call `cache.load_results(key)` then `cache.parse_indices(spec, total)` to resolve user specs (`1,3-5`) into 0-based positions before fetching by ID.

**5. Error handling is centralized.** Use `@handle_graph_errors` on every `run_with_arguments`. Call `raise_if_forbidden(exc)` when a 403 should surface required scope details. Never swallow `ODataError` silently. `ODataError` is imported inside functions in `errors.py` — that is intentional to break a circular import.

**6. SRP, DRY, composition over inheritance.** Each sub-module owns one capability. Pagination, caching, error handling, and token management live in their own utility modules — do not inline those concerns. Prefer composing `GraphContext` services over building class hierarchies.

## Module Authoring Rules

When adding or changing a module:

1. Create a package under `src/msgraphx/modules/<service>/` or add a sub-module to an existing one.
2. Implement `add_arguments(parser, parents)` and `async run_with_arguments(context, args) -> int`.
3. Decorate `run_with_arguments` with `@handle_graph_errors`.
4. Wire the subparser in the parent `__init__.py` and, if top-level, in `cli.py`.
5. Prototype the Graph API call in [Graph Explorer](https://developer.microsoft.com/en-us/graph/graph-explorer) first — note required scopes and exact response shape before writing model access code.
6. Use `pagination.GraphPaginator` or `pagination.collect_all` for any paginated resource.
7. Use `search_entities` from `core/graph_search.py` for full-text search endpoints.

## JSON Output

Every module supports `--json` (global flag). When set:
- `context.json_output` is `True` inside the module.
- Skip all `console.print()` calls — output nothing to stdout except the final JSON.
- At the end of `run_with_arguments`, call `output.print_json(data)` with a list of dicts or a dict.
- Logger continues writing to stderr at normal level — this does not pollute the stdout JSON stream.

Import: `from ...utils import output`, then call `output.print_json(data)`.

The `--save-dir` argument in `aad search` saves results to disk files. It is separate from `--json` (stdout) and can be combined with it.

## Logging and Output Rules

| Level | Use |
|-------|-----|
| `TRACE` | Deep developer inspection — SDK request params, loop iterations, internal values. Not for operators. |
| `DEBUG` | Internal state an operator might need to diagnose their usage |
| `INFO` | Normal operator-visible progress |
| `SUCCESS` | Confirmed completed operation |
| `WARNING` | Recoverable degradation |
| `ERROR` | Recoverable failure — tool continues after logging |

In `except Exception` handlers, choose based on whether the tool stops:
- **Tool stops** (re-raise / fatal): `logger.exception("...")` — ERROR level + full traceback, no need to interpolate `{exc}`.
- **Tool continues** (recoverable / per-item): `logger.error(f"...: {exc}")` — message only, no traceback.

Logs go to stderr (colored) and rotate to `~/.local/state/msgraphx/logs/msgraphx.log`. The operator controls the level via `--trace`, `--debug`, or `--log-level`. Never use `print()` for operational output — use `logger` or the shared `console` (Rich) instance from `utils/console.py`.

## Import Conventions

```python
# Built-in imports
from __future__ import annotations  # always first

import argparse
from pathlib import Path

# External library imports
from loguru import logger
from msgraph.generated.models.drive_item import DriveItem

# Local library imports
from ...core.context import GraphContext
from ...utils import cache, pagination          # module-qualified calls preferred
from ...utils.errors import handle_graph_errors
```

- Three groups, each with its section comment, separated by blank lines.
- Prefer module imports for utilities: `pagination.collect_all(...)`, `cache.save_results(...)`, not bare names.
- `from __future__ import annotations` is the first line in every file.

## Definition of Done

A change is not complete until all are true:

1. The Graph API call was prototyped in Graph Explorer and required scopes are identified.
2. SDK-generated request builders and typed models are used wherever available.
3. `run_with_arguments` is decorated with `@handle_graph_errors`.
4. Paginated resources go through `GraphPaginator` or `pagination.collect_all`.
5. Imports follow the three-group convention with module-qualified utility calls.
6. Relevant tests were updated or added.
7. No unrelated refactors or formatting churn.
