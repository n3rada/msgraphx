# msgraphx/core/context.py

from dataclasses import dataclass
from typing import TYPE_CHECKING, Optional

if TYPE_CHECKING:
    from msgraph import GraphServiceClient
    from msgraph.generated.models.user import User


@dataclass
class GraphContext:
    """
    Shared runtime context for Graph API operations.

    This object is passed to modules to provide access to:
    - The authenticated Graph client
    - Cached user information (for delegated auth)
    - Authentication mode (app-only vs delegated)
    - Other shared runtime state
    """

    graph_client: "GraphServiceClient"
    is_app_only: bool
    region: str = "EMEA"
    cached_user: Optional["User"] = None

    @property
    def is_delegated(self) -> bool:
        """Check if using delegated authentication."""
        return not self.is_app_only
