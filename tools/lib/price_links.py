"""Centralized search-URL templates for manual price cross-checking.

These platforms don't offer public search APIs, and Shopee/Lazada actively
block automated scraping, so this only ever builds links for the user to
open and eyeball themselves -- it never fetches or parses results.
"""

from urllib.parse import quote_plus

SHOPEE_SEARCH_URL_TEMPLATE = "https://shopee.ph/search?keyword={query}"
LAZADA_SEARCH_URL_TEMPLATE = "https://www.lazada.com.ph/catalog/?q={query}"

# This location-less form can redirect to a login/location prompt depending
# on the viewer's browser session. If it stops returning useful results,
# swap in a location-scoped URL instead, e.g.:
#   "https://www.facebook.com/marketplace/manila/search/?query={query}"
FB_MARKETPLACE_SEARCH_URL_TEMPLATE = (
    "https://www.facebook.com/marketplace/search/?query={query}"
)


def build_price_check_links(search_term: str) -> dict[str, str]:
    """Build Shopee/Lazada/Facebook Marketplace search URLs for a term."""
    query = quote_plus(search_term.strip())
    return {
        "Shopee": SHOPEE_SEARCH_URL_TEMPLATE.format(query=query),
        "Lazada": LAZADA_SEARCH_URL_TEMPLATE.format(query=query),
        "Facebook Marketplace": FB_MARKETPLACE_SEARCH_URL_TEMPLATE.format(query=query),
    }
