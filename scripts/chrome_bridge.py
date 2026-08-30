"""
Chrome bridge for the Find Connections (Basic) skill.

WHY THIS EXISTS
---------------
The skill runs as a normal Python process. The Claude-in-Chrome MCP tools
(``navigate`` / ``get_page_text``) can only be invoked by Claude (the agent),
never by a child Python process. So the skill cannot call the browser directly.

This module bridges that gap with a small protocol:

    1. PLAN  — Python enumerates every LinkedIn URL the run will need and prints
               them to stdout. No queue file, no browser, no Excel changes.
    2. FETCH — Claude reads the URL list from stdout, drives Chrome to each URL,
               and stores the page text in a *cache* JSON file
               (``{url: page_text}``) in the ephemeral work dir.
    3. APPLY — Python re-runs the skill with ``CachedFetcher`` as ``fetch_fn``.
               Every fetch is served from the cache; Excel is written normally;
               the cache is deleted once the run applies cleanly.

``CachedFetcher`` records any URL it was asked for but did not find in the cache
(a "miss"), so the orchestrator can fetch the stragglers (e.g. pagination pages
that only become reachable once page 1 is known to have results) and loop until
there are no misses left.
"""

import json
import os
from typing import Dict, List, Optional

from url_parser import LinkedInURLParser, is_company_people_url, paginate_search_url


# --------------------------------------------------------------------------- #
# Cache-backed fetch_fn                                                         #
# --------------------------------------------------------------------------- #

class CachedFetcher:
    """
    A ``fetch_fn`` that serves LinkedIn page text from a JSON cache on disk.

    The cache is a flat ``{url: page_text}`` mapping written by Claude after
    driving Chrome. On a cache miss the fetcher returns ``""`` (the pure parsers
    treat empty text as "no results") and records the URL in ``self.misses`` so
    the caller can fetch it and retry.
    """

    def __init__(self, cache_path: str):
        self.cache_path = cache_path
        self.cache: Dict[str, str] = {}
        self.misses: List[str] = []
        if cache_path and os.path.exists(cache_path):
            with open(cache_path, "r", encoding="utf-8") as f:
                self.cache = json.load(f)

    def __call__(self, url: str) -> str:
        if url in self.cache:
            return self.cache[url] or ""
        if url not in self.misses:
            self.misses.append(url)
        return ""


# --------------------------------------------------------------------------- #
# Offline URL planner                                                          #
# --------------------------------------------------------------------------- #

def plan_fetch_urls(
    ws,
    structure,
    profile_type: str,
    name: str,
    linkedin_url: str,
    parser: Optional[LinkedInURLParser] = None,
    max_pages: int = 5,
) -> List[str]:
    """
    Enumerate every LinkedIn URL a run will fetch — without a browser.

    * ``person`` (adding a contact column): for each company row, that company's
      people-search target. The search URL depends only on the company, not the
      contact, so one target per company covers the whole column — and covers a
      *batch* of contacts added in the same run.
    * ``company`` (adding a company row): the company profile URL (for industry)
      plus that company's people-search target.

    The people-tab surface (``/company/<slug>/people/``) is enumerated as a single
    URL: Claude exhausts its lazy "Show more results" block during FETCH, so
    APPLY drives that expansion directly rather than through ``?page=N``. Only the
    ``/search/results/people/`` keyword fallback is enumerated across pages.

    Returns a de-duplicated, order-preserving list.
    """
    parser = parser or LinkedInURLParser()
    urls: List[str] = []

    def add(u: str) -> None:
        if u and u not in urls:
            urls.append(u)

    def add_search_target(company_name: Optional[str], company_url: Optional[str]) -> None:
        base = parser.build_people_search_url(
            company_name=company_name, company_url=company_url
        )
        if is_company_people_url(base):
            add(base)  # lazy people tab — one URL, expanded live during FETCH
        else:
            for page_num in range(1, max_pages + 1):
                add(paginate_search_url(base, page_num))

    if profile_type == "person":
        for row in range(structure.data_start_row, ws.max_row + 1):
            company_name = ws.cell(row=row, column=structure.company_col).value
            if not company_name:
                continue
            company_url = (
                ws.cell(row=row, column=structure.url_col).value
                if structure.url_col else None
            )
            add_search_target(str(company_name), company_url)

    elif profile_type == "company":
        add(linkedin_url)  # company profile page → industry
        add_search_target(name, linkedin_url)

    return urls


# --------------------------------------------------------------------------- #
# Cache file helpers                                                           #
# --------------------------------------------------------------------------- #
#
# There is deliberately no queue file: ``--plan`` prints the URL list to stdout
# and Claude reads it from there. The only run artifact is the response cache,
# which lives in the ephemeral work dir (see ``workspace.py``) and is deleted
# once a run applies cleanly.

def load_cache(cache_path: str) -> Dict[str, str]:
    """Load the ``{url: page_text}`` cache, or an empty dict if absent."""
    if not cache_path or not os.path.exists(cache_path):
        return {}
    with open(cache_path, "r", encoding="utf-8") as f:
        return json.load(f)


def save_cache(cache_path: str, cache: Dict[str, str]) -> None:
    """Persist the ``{url: page_text}`` cache to disk."""
    with open(cache_path, "w", encoding="utf-8") as f:
        json.dump(cache, f, indent=2, ensure_ascii=False)


def update_cache(cache_path: str, url: str, page_text: str) -> Dict[str, str]:
    """Add/overwrite one URL's page text and persist. Returns the updated cache."""
    cache = load_cache(cache_path)
    cache[url] = page_text
    save_cache(cache_path, cache)
    return cache
