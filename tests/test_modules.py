"""Tests for module discovery and imports."""

from __future__ import annotations

import importlib
import pkgutil

import msgraphx.modules.outlook as outlook
import msgraphx.modules.sharepoint as sharepoint
import msgraphx.modules.aad as aad
import msgraphx.modules.me as me


class TestModuleDiscovery:
    """Ensure all module packages are importable and have expected structure."""

    def test_outlook_contacts_importable(self):
        mod = importlib.import_module("msgraphx.modules.outlook.contacts")
        assert hasattr(mod, "add_arguments")
        assert hasattr(mod, "run_with_arguments")

    def test_sharepoint_search_importable(self):
        mod = importlib.import_module("msgraphx.modules.sharepoint.search")
        assert hasattr(mod, "add_arguments")
        assert hasattr(mod, "run_with_arguments")

    def test_sharepoint_download_importable(self):
        mod = importlib.import_module("msgraphx.modules.sharepoint.download")
        assert hasattr(mod, "add_arguments")
        assert hasattr(mod, "run_with_arguments")

    def test_aad_search_importable(self):
        mod = importlib.import_module("msgraphx.modules.aad.search")
        assert hasattr(mod, "add_arguments")
        assert hasattr(mod, "run_with_arguments")

    def test_me_groups_importable(self):
        mod = importlib.import_module("msgraphx.modules.me.groups")
        assert hasattr(mod, "add_arguments")
        assert hasattr(mod, "run_with_arguments")

    def test_pkgutil_discovers_submodules(self):
        """Verify pkgutil can discover submodules like the CLI does."""
        for package in (outlook, sharepoint, aad, me):
            modules = list(
                pkgutil.iter_modules(package.__path__, package.__name__ + ".")
            )
            # Each package should have at least one discoverable submodule
            assert len(modules) >= 1, f"{package.__name__} has no submodules"


class TestCoreImports:
    """Ensure core modules import without circular dependency issues."""

    def test_context_importable(self):
        from msgraphx.core.context import GraphContext

        assert GraphContext is not None

    def test_graph_search_importable(self):
        from msgraphx.core.graph_search import SearchOptions, search_entities

        assert SearchOptions is not None
        assert search_entities is not None
