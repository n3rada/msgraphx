# msgraphX: Copilot Instructions

This file is the canonical AI guidance for this repository.

## Context

msgraphX is a red team / penetration testing tool that exploits the Microsoft Graph API to harvest SharePoint files, Outlook mail, Teams messages, and Microsoft 365 data during offensive operations. It is in active beta — breaking changes are expected and backward compatibility is never a concern.

`verify=False` on HTTP calls is intentional (intercepting proxy / self-signed cert context). Never flag it.

## Read Order

1. [README.md](../README.md) — CLI behavior, authentication modes, all module usage.
2. [DEVELOPMENT.md](../DEVELOPMENT.md) — architecture, design principles, SDK rules, and extension model (source of truth).
3. [src/msgraphx/cli.py](../src/msgraphx/cli.py) — runtime flow, auth bootstrap, argument dispatch.
4. [tests/](../tests/) — expected behavior and regression boundaries.

## Source Map

| Path | Purpose |
|------|---------|
| `src/msgraphx/cli.py` | Entry point, argument dispatch, auth bootstrap |
| `src/msgraphx/core/context.py` | `GraphContext` — shared runtime state passed to all modules |
| `src/msgraphx/core/graph_search.py` | `POST /search/query` wrapper (`SearchOptions`, `search_entities`) |
| `src/msgraphx/modules/` | Per-service capability packages (SharePoint, Outlook, Teams, AAD, Me) |
| `src/msgraphx/utils/pagination.py` | `GraphPaginator`, `collect_all` — all paginated SDK calls go here |
| `src/msgraphx/utils/errors.py` | `handle_graph_errors`, `AuthenticationError`, `ForbiddenGraphError` |
| `src/msgraphx/utils/cache.py` | XDG result cache, `parse_indices` |
| `src/msgraphx/utils/tokens.py` | Token loading, refresh, `.roadtools_auth` interop |
| `src/msgraphx/utils/logbook.py` | Loguru setup, log rotation |
| `src/msgraphx/utils/console.py` | Shared Rich `Console` instance (stdout) |
| `src/msgraphx/utils/output.py` | `print_json(data)` — JSON-to-stdout for `--json` mode |

## Architecture and Design Principles

Detailed rationale lives in [DEVELOPMENT.md](../DEVELOPMENT.md). If this file and DEVELOPMENT.md diverge, follow DEVELOPMENT.md.

**GraphContext is the only shared state.** Modules receive a `GraphContext` and must not construct their own `GraphServiceClient`. Never bypass it.

**Module contract is fixed.** Every module exposes exactly `add_arguments(parser, parents)` and `async run_with_arguments(context, args) -> int`, decorated with `@handle_graph_errors`.

**Two distinct API paths:**
- `core/graph_search.py` (`search_entities`) for full-text search (SharePoint `DriveItem`, Teams `ChatMessage`, Outlook `Message`). Handles pagination and the Exchange 1 000-hit cap.
- SDK list builders (`client.users`, `client.groups`, etc.) for enumeration, always via `pagination.GraphPaginator` or `pagination.collect_all`. Never hand-roll `@odata.nextLink` loops.

**Cache → index → fetch is the download pattern.** Search modules write results via `cache.save_results`. Download modules call `cache.load_results` then `cache.parse_indices` to resolve user index specs before fetching by ID.

**Error handling is centralized.** `@handle_graph_errors` on every `run_with_arguments`. `raise_if_forbidden(exc)` for 403s. Never swallow `ODataError` silently.

**SRP, DRY, composition.** Each sub-module owns one capability. Pagination, caching, error handling, and token management live in their own utility modules — do not inline those concerns.

## Module Authoring Rules

1. Create a package under `src/msgraphx/modules/<service>/` or add a sub-module to an existing one.
2. Implement `add_arguments` and `async run_with_arguments`, decorated with `@handle_graph_errors`.
3. Wire the subparser in the parent `__init__.py` and, if top-level, in `cli.py`.
4. Prototype the Graph API call in [Graph Explorer](https://developer.microsoft.com/en-us/graph/graph-explorer) first.
5. Use `pagination.GraphPaginator` or `pagination.collect_all` for paginated resources.
6. Use `search_entities` from `core/graph_search.py` for full-text search endpoints.

## JSON Output

Every module supports the global `--json` flag. When `context.json_output` is `True`:
- Suppress all `console.print()` calls — nothing goes to stdout except the final JSON.
- Call `output.print_json(data)` at the end of `run_with_arguments` with a list of dicts or a dict.
- Log level is automatically raised to `WARNING` (stderr) to keep the stdout stream clean.
- Import: `from ...utils import output`, call: `output.print_json(data)`.

The `--save-dir` flag in `aad search` saves to disk files — separate from `--json` (stdout).

## Logging Levels

| Level | Use |
|-------|-----|
| `TRACE` | Deep developer inspection — SDK request params, loop iterations, internal values |
| `DEBUG` | Internal state useful to diagnose operator usage |
| `INFO` | Normal operator-visible progress |
| `SUCCESS` | Confirmed completed operation |
| `WARNING` | Recoverable degradation |
| `ERROR` | Non-recoverable failure that does not halt the process |

Never use `print()` for operational output. Use `logger` (loguru) or the shared `console` (Rich) from `utils/console.py`.

## Import Conventions

- Three groups in order: `# Built-in imports`, `# External library imports`, `# Local library imports`, separated by blank lines.
- `from __future__ import annotations` is always the first line.
- Prefer module-qualified utility calls: `pagination.collect_all(...)`, `cache.save_results(...)` — not bare names.

## Python Rules

- Python 3.12+. Use `X | Y` and `X | None` union syntax throughout.
- Annotate all function signatures.
- Avoid broad `except Exception` unless re-raising or translating to a domain error.
- In `except Exception` handlers: if the tool stops (re-raise / fatal), use `logger.exception("...")` (ERROR + traceback, no `{exc}` interpolation needed). If the tool continues (recoverable), use `logger.error(f"...: {exc}")`.
- All Graph SDK calls are async. Use `asyncio.gather()` for independent concurrent calls.

## Definition of Done

1. Graph API call prototyped in Graph Explorer; required scopes identified.
2. SDK request builders and typed models used wherever available.
3. `run_with_arguments` decorated with `@handle_graph_errors`.
4. Paginated resources go through `GraphPaginator` or `pagination.collect_all`.
5. Three-group import convention followed with module-qualified utility calls.
6. Relevant tests updated or added.
7. No unrelated refactors or formatting churn.
