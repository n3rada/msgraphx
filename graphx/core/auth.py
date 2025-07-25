# Built-in imports
import asyncio
import argparse
import json
import secrets
from urllib.parse import urlencode, parse_qs, urlparse


# External library imports
import pkce
import httpx
from loguru import logger

from playwright.async_api import async_playwright


GRAPH_EXPLORER_APP_ID = "de8bc8b5-d9f9-48b1-a8ad-b748da725064"
REDIRECT_URI = "https://developer.microsoft.com/en-us/graph/graph-explorer"

SCOPE = "openid https://graph.microsoft.com/.default offline_access"


async def run_with_arguments(args: argparse.Namespace) -> int:
    tokens = await get_tokens(prt_cookie=args.prt_cookie, headless=args.headless)

    if not tokens:
        return 1

    print(tokens["access_token"])

    with open(".roadtools_auth", mode="w", encoding="utf-8") as file_obj:
        json.dump(
            {
                "accessToken": tokens["access_token"],
                "refreshToken": tokens["refresh_token"],
                "expiresIn": tokens["expires_in"],
            },
            file_obj,
            indent=4,
        )

    logger.success("✅ Tokens saved to .roadtools_auth.")
    return 0


def add_arguments(parser: argparse.ArgumentParser):
    parser.add_argument(
        "--prt-cookie",
        type=str,
        default=None,
        help="X-Ms-RefreshTokenCredential PRT cookie",
    )
    parser.add_argument(
        "--headless",
        action="store_true",
        default=False,
        help="Run the browser in headless mode (default: False)",
    )


def standalone() -> int:
    parser = argparse.ArgumentParser(
        prog="graphx-auth",
        add_help=True,
        description="Microsoft Graph Explorer authentication module.",
    )

    add_arguments(parser)
    args = parser.parse_args()
    try:
        return asyncio.run(run_with_arguments(args))
    except KeyboardInterrupt:
        logger.warning("🛑 Interrupted by user.")
        return 0


# Core function to get tokens using Playwright


async def get_tokens(prt_cookie: str = None, headless: bool = False) -> dict:
    response_dict = {"refresh_token": None, "access_token": None, "expires_in": None}
    code_verifier, code_challenge = pkce.generate_pkce_pair()
    state = secrets.token_urlsafe(32)

    params = {
        "client_id": GRAPH_EXPLORER_APP_ID,
        "scope": SCOPE,
        "redirect_uri": REDIRECT_URI,
        "response_type": "code",
        "code_challenge": code_challenge,
        "code_challenge_method": "S256",
        "state": state,
    }

    auth_url = f"https://login.microsoftonline.com/common/oauth2/v2.0/authorize?{urlencode(params)}"

    logger.info("🔐 Starting authentication process using Playwright (async)")

    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=headless, args=["--no-sandbox"])
        context = await browser.new_context()

        if prt_cookie:
            await context.add_cookies(
                [
                    {
                        "name": "x-ms-RefreshTokenCredential",
                        "value": prt_cookie,
                        "domain": "login.microsoftonline.com",
                        "path": "/",
                        "httpOnly": True,
                        "secure": True,
                    }
                ]
            )

        page = await context.new_page()
        logger.info(f"🔗 Opening auth URL: {auth_url}")

        await page.goto(auth_url)
        await page.wait_for_load_state("load")

        logger.info("🔍 Waiting for authentication to complete")

        try:
            await page.wait_for_url(REDIRECT_URI + "*", timeout=2 * 60 * 1000)
        except TimeoutError:
            logger.error("⏱️ Timeout waiting for auth redirect.")
            return None
        except Exception as exc:
            logger.error(f"❌ Error during auth redirect: {exc}")
            return None
        else:
            final_url = page.url
            logger.success("🔄 Redirection received.")
        finally:
            await context.close()
            await browser.close()
            logger.info("🖥️ Browser closed.")

    code = parse_qs(urlparse(final_url).query).get("code", [None])[0]

    if not code:
        logger.error("❌ Authorization code not found in redirect URL.")
        return None

    logger.info("🔑 Exchanging authorization code for tokens")

    async with httpx.AsyncClient() as client:
        response = await client.post(
            "https://login.microsoftonline.com/common/oauth2/v2.0/token",
            data={
                "client_id": GRAPH_EXPLORER_APP_ID,
                "redirect_uri": REDIRECT_URI,
                "scope": SCOPE,
                "code": code,
                "code_verifier": code_verifier,
                "grant_type": "authorization_code",
                "claims": '{"access_token":{"xms_cc":{"values":["CP1"]}}}',
            },
            headers={"Origin": "https://developer.microsoft.com"},
        )

    if response.status_code != 200:
        logger.error(f"❌ Token exchange failed: {response.text}")
        return None

    logger.success("✅ Token exchange successful")

    tokens = response.json()
    response_dict["refresh_token"] = tokens.get("refresh_token")
    response_dict["access_token"] = tokens.get("access_token")
    response_dict["expires_in"] = tokens.get("expires_in")

    return response_dict
