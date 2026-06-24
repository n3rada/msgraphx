# msgraphX: Development Guide

This file is the canonical development guidance for this repository.

## Runtime Requirement

- Python version baseline is 3.12+ (see [pyproject.toml](pyproject.toml): `requires-python = ">=3.12,<4.0"`).
- Do not introduce language features or dependencies incompatible with Python 3.12.
- Use `uv` for environment and dependency workflows in this repo.

## Read Order

1. [README.md](README.md) — CLI behavior, authentication modes, supported modules.
2. [src/msgraphx/cli.py](src/msgraphx/cli.py) — runtime flow and top-level argument handling.
3. [src/msgraphx/core/context.py](src/msgraphx/core/context.py) — shared runtime state passed to all modules.
4. [src/msgraphx/utils/](src/msgraphx/utils/) — pagination, error handling, caching, token management.
5. [tests/](tests/) — expected behavior and regression boundaries.

---

## Core Rule: Use the Microsoft Graph Python SDK

This project uses the official [`msgraph-sdk`](https://github.com/microsoftgraph/msgraph-sdk-python) (Microsoft Graph Python SDK). **Always use the SDK for any Graph API call. Never use raw `httpx` to call a Graph endpoint.**

- **Always target the latest SDK.** When a new Graph API feature or endpoint becomes available in the SDK, use it. Do not implement raw `httpx` calls to endpoints already exposed by the SDK.
- Use generated request builders (`client.users`, `client.groups`, `client.drives`, etc.) and their typed query parameter classes. Never construct Graph API URLs by hand when a builder exists.
- Use the SDK's typed models (e.g., `DriveItem`, `Message`, `User`) rather than parsing raw JSON dicts.
- For paginated resources, use [`GraphPaginator`](src/msgraphx/utils/pagination.py) or `pagination.collect_all()` — never implement ad-hoc `@odata.nextLink` loops.
- For parallel calls to independent endpoints, use `asyncio.gather()` on SDK awaitable calls — not a hand-rolled `$batch` POST via httpx.
- For batch requests, use `context.graph_client.batch()` (`BatchRequestBuilder` from `msgraph_core`) — not a raw `POST /$batch` via httpx.

### Batch requests via SDK

```python
from msgraph_core.requests.batch_request_content import BatchRequestContent

batch_content = BatchRequestContent()
batch_content.add_request_information(
    context.graph_client.users.by_user_id(uid).to_get_request_information(config),
    request_id="user",
)
batch_content.add_request_information(
    context.graph_client.users.by_user_id(uid).owned_objects.to_get_request_information(),
    request_id="ownedObjects",
)

batch_resp = await context.graph_client.batch().post(batch_content)

status_codes = batch_resp.get_response_status_codes()
item = batch_resp.get_response_by_id("user")         # BatchResponseItem
body = json.loads(item.body.read()) if item and item.body else {}
```

### When httpx is allowed

`httpx` is allowed **only** for non-Graph endpoints that the SDK does not cover:

| Module | Endpoint | Reason |
|--------|----------|--------|
| `me/drive.py` (`_chunked_upload`) | Azure Blob upload session URL | Upload session URLs are not Graph endpoints |
| `cli.py` (`_check_public_ip`) | `api.ipify.org` | External IP check, not a Graph endpoint |

Any other use of `httpx` to call a `graph.microsoft.com` URL is a violation of this rule and must be refactored to use the SDK.

### Testing Graph API calls

Use [Microsoft Graph Explorer](https://developer.microsoft.com/en-us/graph/graph-explorer) to prototype and validate Graph API queries before implementing them in code. Graph Explorer lets you:

- Test endpoint availability and required scopes for a given token.
- Inspect the exact response shape before writing model access code.
- Verify query parameters (`$filter`, `$select`, `$top`, `$search`) without a full auth flow.

---

## Architecture

### GraphContext

[`GraphContext`](src/msgraphx/core/context.py) is the single shared runtime object passed to every module. It carries:

- `graph_client` — the authenticated `GraphServiceClient` instance.
- `is_app_only` — whether authentication is app-only or delegated.
- `cached_user` — the authenticated user's `User` object (delegated only).
- `token_scopes` — frozenset of granted scopes for delegated tokens.

Modules receive `GraphContext` and must not construct their own `GraphServiceClient`.

### Module structure

Each top-level capability (SharePoint, Outlook, Teams, AAD, Me) is a package under [`src/msgraphx/modules/`](src/msgraphx/modules/). Each package exposes exactly two public functions consumed by the CLI:

```python
def add_arguments(parser: argparse.ArgumentParser, parents: list | None = None) -> None: ...

async def run_with_arguments(context: GraphContext, args: argparse.Namespace) -> int: ...
```

Sub-capabilities (e.g., `teams/chat.py`, `teams/channel.py`) follow the same two-function contract. The package `__init__.py` dispatches to them.

### Error handling

- Decorate all `run_with_arguments` entry points with `@handle_graph_errors` from [`utils/errors.py`](src/msgraphx/utils/errors.py). This converts `InvalidAuthenticationToken` errors into `AuthenticationError` and exits cleanly.
- Raise `ForbiddenGraphError` (via `raise_if_forbidden()`) when a 403 is received and the caller can make the missing scope actionable to the operator.
- Do not swallow `ODataError` silently; at minimum log the error code and message.

### Pagination

Always use the pagination utilities instead of manual `@odata.nextLink` loops:

```python
# Preferred: async iteration
async for user in pagination.GraphPaginator(context.graph_client.users, request_config):
    ...

# Preferred: collect all pages at once
results = await pagination.collect_all(context.graph_client.groups, request_config)
```

---

## Python Rules

### 1. Imports

- Group imports in this order, with an explicit section comment for each group:
  1. `# Built-in imports` — standard library
  2. `# External library imports` — third-party packages
  3. `# Local library imports` — intra-package imports
- Separate groups with a blank line. No blank lines within a group.
- Prefer module imports for utility modules over importing individual functions. Call them as `pagination.collect_all(...)`, `cache.save_results(...)`, not bare `collect_all(...)`. This makes the call site self-documenting and avoids name collisions.
- Keep imports at module top. Use a local import only to break a genuine circular dependency or to defer an expensive optional import.
- Always include `from __future__ import annotations` as the first import in every file.

Example of correct import block:

```python
# Built-in imports
from __future__ import annotations

import argparse
import json
from pathlib import Path

# External library imports
from loguru import logger
from msgraph.generated.models.drive_item import DriveItem
from rich.table import Table

# Local library imports
from ...core.context import GraphContext
from ...utils import cache, pagination
from ...utils.errors import handle_graph_errors
```

### 2. Typing and modern Python

- Use modern Python 3.12 union syntax: `X | Y` and `X | None` instead of `Optional[X]` or `Union[X, Y]`.
- Annotate all function signatures (parameters and return types). Use `None` as the return type for functions that only produce side effects.
- Prefer concrete types. Avoid `Any` except where the SDK's own types are unresolvable at annotation time.
- Use `frozenset`, `tuple`, and other immutable types where mutation is not intended.

### 3. JSON output

All modules support the global `--json` flag. When `context.json_output` is `True`:

- Do not call `console.print()` — output nothing to stdout except the final JSON.
- At the end of `run_with_arguments`, call `output.print_json(data)` where `data` is a list of dicts or a dict.
- Logger continues writing to stderr at normal level — stdout stays clean for the JSON consumer.
- Import via `from ...utils import output`, call as `output.print_json(data)`.

```python
if context.json_output:
    output.print_json([{"id": item.id, "name": item.name, ...} for item in results])
    return 0
# existing Rich rendering below
```

The `--save-dir` argument in `aad search` saves results to disk as JSON files — this is independent of `--json` (stdout) and both can be used simultaneously.

### 4. Error handling and logging

- Avoid broad `except Exception` unless you are re-raising, translating to a domain error, or intentionally degrading behavior.
- In `except Exception` handlers, choose based on whether the tool stops or continues:
  - **Tool stops** (re-raise / fatal failure): `logger.exception("...")` — ERROR level with full traceback, no need to interpolate `{exc}` in the message.
  - **Tool continues** (recoverable / per-item failure): `logger.error(f"...: {exc}")` — message only, no traceback.
- Use `logger.trace(...)` for deep developer-facing logs (request parameters, intermediate values, per-iteration data); `logger.debug(...)` for general internal state; `logger.success(...)` for completed operations; `logger.warning(...)` for recoverable issues; `logger.error(...)` for non-recoverable ones that do not halt the process.
- Never use `except Exception: pass` — silent swallowing masks real failures.

### 4. Async

- All Graph SDK calls are async. `run_with_arguments` and any function that calls `graph_client.*` must be `async def`.
- Use `asyncio.gather()` for concurrent independent Graph API calls (e.g., fetching multiple drives in parallel).
- Do not block the event loop with synchronous I/O inside async functions; use `asyncio.to_thread()` for file operations when needed.

### 5. Code hygiene

- Write comments only when the *why* is non-obvious. Do not narrate what the code does.
- Keep diffs minimal. Do not reformat or reorganize code unrelated to the change being made.
- Do not add `__all__` unless a module is intended to be a public library surface.
- Magic numbers used as Graph API page sizes, limits, or thresholds must be assigned to a named constant at module level.

---

## Design Principles

### Single Responsibility Principle (SRP)

Each module, class, and function should have one reason to change.

- Each sub-module (`chat.py`, `download.py`, `search.py`, …) is responsible for exactly one capability. Do not add unrelated logic because it is convenient.
- `GraphContext` carries shared runtime state; it does not execute queries or format output.
- Formatting and display logic belongs in the module that owns the output, not in utilities.
- Pagination, caching, error handling, and token management each live in their own utility module. Do not inline those concerns into capability modules.

### Open/Closed Principle

The module system is open for extension, closed for modification.

- Add new capabilities by creating a new sub-module and wiring it into the parent `__init__.py` and `cli.py`. Do not change the routing logic of existing modules to accommodate a new one.
- The `add_arguments` / `run_with_arguments` contract is the stable interface. New modules implement it; they do not change it.

### DRY (Don't Repeat Yourself)

- Pagination logic lives in `utils/pagination.py`. If you find yourself writing a `while odata_next_link` loop in a module, move it there.
- ODataError inspection and auth-error detection live in `utils/errors.py`. Do not duplicate that logic in individual modules.
- If two modules need the same Graph query or data-shaping step, extract it to a shared helper — either in the relevant `utils/` module or in a `core/` helper — rather than copying it.

### Composition over Inheritance

Prefer composing small, focused functions and passing `GraphContext` over building class hierarchies.

- Modules are collections of functions, not class trees. A module that needs helper state should use a local dataclass or a plain function with parameters, not a base class with overrides.
- `GraphContext` is a composition point: it carries the client, auth mode, and cached state without encoding behavior. Modules receive it; they do not subclass it.

### Fail Fast and Observable

- Surface errors as early and specifically as possible. A precise `ForbiddenGraphError` with the required scope named is more useful than a generic log line.
- Log at the right level: `logger.trace` for deep developer inspection (request parameters, intermediate values, loop iterations), `logger.debug` for internal state, `logger.info` for operator-visible progress, `logger.warning` for recoverable degradation, `logger.error` for non-recoverable failures, `logger.success` for completed operations.
- Do not hide failures behind a default return value. If a function cannot complete its contract, raise or return `None` explicitly and let the caller decide.

### KISS (Keep It Simple)

- Solve the problem at hand. Do not design for hypothetical future requirements.
- Three similar lines are better than a premature abstraction. Extract only when the duplication has already appeared at least twice and a third occurrence is likely.
- Prefer a single readable function over a clever pipeline that requires reading five layers of indirection to understand.

---

## Source Map

| Path | Purpose |
|------|---------|
| [`src/msgraphx/cli.py`](src/msgraphx/cli.py) | Entry point, argument dispatch, auth bootstrap |
| [`src/msgraphx/core/context.py`](src/msgraphx/core/context.py) | `GraphContext` dataclass |
| [`src/msgraphx/core/graph_search.py`](src/msgraphx/core/graph_search.py) | Cross-entity Graph Search API wrapper |
| [`src/msgraphx/modules/`](src/msgraphx/modules/) | Per-service capability modules |
| [`src/msgraphx/utils/pagination.py`](src/msgraphx/utils/pagination.py) | `GraphPaginator` and `collect_all` |
| [`src/msgraphx/utils/errors.py`](src/msgraphx/utils/errors.py) | `handle_graph_errors`, `AuthenticationError`, `ForbiddenGraphError` |
| [`src/msgraphx/utils/cache.py`](src/msgraphx/utils/cache.py) | Result persistence and retrieval |
| [`src/msgraphx/utils/tokens.py`](src/msgraphx/utils/tokens.py) | Token loading, refresh, `.roadtools_auth` interop |
| [`src/msgraphx/utils/logbook.py`](src/msgraphx/utils/logbook.py) | Loguru configuration |
| [`src/msgraphx/utils/console.py`](src/msgraphx/utils/console.py) | Shared Rich `Console` instance (stdout) |
| [`src/msgraphx/utils/output.py`](src/msgraphx/utils/output.py) | `print_json(data)` — JSON-to-stdout for `--json` mode |
| [`tests/`](tests/) | pytest suite |

---

## Adding a New Module

1. Create a package under the appropriate parent (or a new one at `src/msgraphx/modules/<service>/`).
2. Implement `add_arguments(parser, parents)` and `async run_with_arguments(context, args) -> int`.
3. Decorate `run_with_arguments` with `@handle_graph_errors`.
4. Wire the subparser in the parent `__init__.py` and in `cli.py`.
5. Use `pagination.GraphPaginator` or `pagination.collect_all` for any paginated resource.
6. Prototype the Graph API call in [Graph Explorer](https://developer.microsoft.com/en-us/graph/graph-explorer) first and note the required scopes.

---

## Git Conventions

- Commit messages must follow [Conventional Commits](https://www.conventionalcommits.org/): `type(scope): short description`. One-liner, no body.
- Always sign commits with `-S`.

## Testing

- Run the suite with `uv run pytest -v`.
- Tests are in [`tests/`](tests/) and cover utilities (cache, dates, tokens) and module smoke tests.
- Keep test output stable; avoid tests that depend on network access or live Graph API tokens.

## Definition of Done

A change is complete only when all are true:

1. The feature was prototyped or validated in Graph Explorer before implementation.
2. The SDK's generated request builders and typed models are used wherever available.
3. All `run_with_arguments` entry points are decorated with `@handle_graph_errors`.
4. Paginated resources go through `GraphPaginator` or `pagination.collect_all`.
5. Imports follow the three-group convention with module-level references for utilities.
6. Relevant tests were updated or added.
