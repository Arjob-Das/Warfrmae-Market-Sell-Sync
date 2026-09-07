"""
Warframe Market Sell Sync - API Integration & Catalog Cache
============================================================
Handles authenticated & public order fetching from Warframe.market,
JWT validation, and items catalog caching.
"""

import os
import json
import time
import base64
import requests
from typing import Dict, List, Tuple

from .config import (
    API_BASE_URL, BROWSER_HEADERS, CACHE_FILE
)
from .sheet_layout import prompt_jwt_token


def get_auth_headers_and_cookies(jwt_token: str) -> Tuple[dict, dict]:
    """Builds Authorization header and cookies from JWT token."""
    headers = BROWSER_HEADERS.copy()
    headers["Authorization"] = f"Bearer {jwt_token}"

    csrf_token = None
    try:
        parts = jwt_token.split(".")
        if len(parts) >= 2:
            payload_b64 = parts[1] + "=" * ((4 - len(parts[1]) % 4) % 4)
            payload_data = json.loads(base64.urlsafe_b64decode(payload_b64).decode("utf-8"))
            csrf_token = payload_data.get("csrf_token")
    except Exception:
        pass

    if csrf_token:
        headers["x-csrftoken"] = csrf_token

    cookies = {"JWT": jwt_token}
    return headers, cookies


def is_jwt_expired(jwt_token: str) -> bool:
    """Checks whether the JWT token timestamp is expired."""
    if not jwt_token:
        return True
    try:
        parts = jwt_token.strip().split(".")
        if len(parts) >= 2:
            payload_b64 = parts[1] + "=" * ((4 - len(parts[1]) % 4) % 4)
            payload_data = json.loads(base64.urlsafe_b64decode(payload_b64).decode("utf-8"))
            exp = payload_data.get("exp")
            if exp and time.time() >= exp:
                return True
            return False
    except Exception:
        pass
    return True


def validate_jwt_token(jwt_token: str) -> bool:
    """Validates whether JWT token can successfully authenticate against the Warframe.market API."""
    if not jwt_token or is_jwt_expired(jwt_token):
        return False
    try:
        headers, cookies = get_auth_headers_and_cookies(jwt_token)
        resp = requests.get(f"{API_BASE_URL}/orders/my", headers=headers, cookies=cookies, timeout=8)
        return resp.status_code in (200, 201)
    except Exception:
        return False


def load_items_catalog() -> Dict[str, dict]:
    """Loads cached items catalog or fetches fresh if missing."""
    if os.path.exists(CACHE_FILE):
        try:
            with open(CACHE_FILE, "r", encoding="utf-8") as f:
                cat = json.load(f)
                if cat and isinstance(cat, dict):
                    return cat
        except Exception:
            pass

    print("[*] Downloading items catalog from warframe.market...")
    try:
        resp = requests.get(f"{API_BASE_URL}/items", headers=BROWSER_HEADERS, timeout=15)
        resp.raise_for_status()
        items_list = resp.json().get("data", [])

        catalog = {}
        for it in items_list:
            i_id = it.get("id")
            en_name = it.get("i18n", {}).get("en", {}).get("name")
            if i_id and en_name:
                catalog[i_id] = {
                    "name": en_name,
                    "slug": it.get("slug", ""),
                    "tags": it.get("tags", []),
                    "maxRank": it.get("maxRank")
                }

        with open(CACHE_FILE, "w", encoding="utf-8") as f:
            json.dump(catalog, f)
        print(f"[+] Cached {len(catalog)} items.")
        return catalog
    except Exception as e:
        print(f"[!] Warning: Could not download items catalog: {e}")
        return {}


def fetch_orders(username: str, jwt_token: str, excel_file: str) -> List[dict]:
    """
    Fetches sell orders from warframe.market.
    Uses authenticated endpoint `/v2/orders/my` if JWT is available,
    otherwise uses public endpoint `/v2/orders/user/{username}`.
    """
    catalog = load_items_catalog()

    if jwt_token:
        headers, cookies = get_auth_headers_and_cookies(jwt_token)
        try:
            resp = requests.get(f"{API_BASE_URL}/orders/my", headers=headers, cookies=cookies, timeout=12)
            if resp.status_code in (200, 201):
                raw_orders = resp.json().get("data", [])
                orders = []
                for o in raw_orders:
                    if o.get("order_type") == "sell" or o.get("type") == "sell":
                        i_id = o.get("itemId") or o.get("item_id")
                        cat_item = catalog.get(i_id, {})
                        base_name = cat_item.get("name") or o.get("item", {}).get("slug", "Unknown Item")
                        rank = o.get("rank")
                        max_rank = cat_item.get("maxRank")
                        if rank is not None and max_rank and max_rank > 0:
                            display_name = f"{base_name} (Rank {rank})"
                        else:
                            display_name = base_name
                        orders.append({
                            "order_id": o.get("id"),
                            "item_id": i_id,
                            "name": display_name,
                            "base_name": base_name,
                            "price": int(o.get("platinum", 0)),
                            "quantity": int(o.get("quantity", 1)),
                            "rank": rank,
                            "visible": o.get("visible", True)
                        })
                return orders
            elif resp.status_code == 401:
                print("[!] Current JWT token has expired or is invalid.")
                fresh_token = prompt_jwt_token(excel_file)
                if fresh_token:
                    return fetch_orders(username, fresh_token, excel_file)
        except Exception as e:
            print(f"[!] Error with authenticated fetch: {e}")

    # Fallback to public orders by username
    if username:
        try:
            print(f"[*] Fetching public sell orders for user '{username}'...")
            resp = requests.get(f"{API_BASE_URL}/orders/user/{username}", headers=BROWSER_HEADERS, timeout=12)
            if resp.status_code == 200:
                raw_orders = resp.json().get("data", [])
                orders = []
                for o in raw_orders:
                    if o.get("order_type") == "sell" or o.get("type") == "sell":
                        i_id = o.get("itemId") or o.get("item_id")
                        cat_item = catalog.get(i_id, {})
                        base_name = cat_item.get("name") or o.get("item", {}).get("slug", "Unknown Item")
                        rank = o.get("rank")
                        max_rank = cat_item.get("maxRank")
                        if rank is not None and max_rank and max_rank > 0:
                            display_name = f"{base_name} (Rank {rank})"
                        else:
                            display_name = base_name
                        orders.append({
                            "order_id": o.get("id"),
                            "item_id": i_id,
                            "name": display_name,
                            "base_name": base_name,
                            "price": int(o.get("platinum", 0)),
                            "quantity": int(o.get("quantity", 1)),
                            "rank": rank,
                            "visible": o.get("visible", True)
                        })
                return orders
        except Exception as e:
            print(f"[!] Error with public fetch: {e}")

    return []
