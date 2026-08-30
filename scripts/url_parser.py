"""
LinkedIn URL parsing utilities.
"""

import re
from typing import Optional, Tuple


def paginate_search_url(base_url: str, page_num: int) -> str:
    """
    Append a page number to a LinkedIn search/people URL.

    Page 1 returns the base URL unchanged. Higher pages append ``page=N`` using
    ``?`` or ``&`` depending on whether the base already carries query params.

    Shared by the live fetcher (``search_company_employees``) and the offline
    planner (``chrome_bridge.plan_fetch_urls``) so both build identical URLs and
    the response cache always hits.
    """
    if page_num <= 1:
        return base_url
    sep = "&" if "?" in base_url else "?"
    return f"{base_url}{sep}page={page_num}"


def is_company_people_url(url: str) -> bool:
    """
    True for a company people-tab URL (``/company/<slug>/people/``).

    This surface loads mutual-connection cards lazily via "Show more results"
    (exhausted by Claude during FETCH), so it is fetched once and never paginated
    with ``?page=N`` — unlike the ``/search/results/people/`` fallback, which is
    genuinely paginated. Both the live fetcher and the offline planner branch on
    this so PLAN and APPLY agree on the URL set and the response cache hits.
    """
    if not url:
        return False
    return "/company/" in url.lower() and "/people/" in url.lower()


class LinkedInURLParser:
    """Parses and normalises LinkedIn URLs."""

    def parse(self, url: str) -> Tuple[Optional[str], Optional[str]]:
        """
        Return (profile_type, profile_id) where profile_type is 'person' or 'company'.
        Returns (None, None) for unrecognised URLs.
        """
        if not url:
            return None, None

        url = url.strip()
        if not url.startswith(("http://", "https://")):
            url = f"https://{url}"

        # Drop trailing slash and query params/fragments
        url = re.sub(r"[?#].*$", "", url).rstrip("/")

        m = re.search(r"linkedin\.com/in/([a-zA-Z0-9\-_%]+)", url, re.IGNORECASE)
        if m:
            return "person", m.group(1)

        m = re.search(r"linkedin\.com/company/([a-zA-Z0-9\-_%]+)", url, re.IGNORECASE)
        if m:
            return "company", m.group(1)

        return None, None

    def normalize(self, url: str) -> Optional[str]:
        """Return canonical https://www.linkedin.com/... URL."""
        profile_type, profile_id = self.parse(url)
        if profile_type == "person":
            return self.build_person_url(profile_id)
        elif profile_type == "company":
            return self.build_company_url(profile_id)
        return None

    @staticmethod
    def build_person_url(profile_id: str) -> str:
        return f"https://www.linkedin.com/in/{profile_id}"

    @staticmethod
    def build_company_url(profile_id: str) -> str:
        return f"https://www.linkedin.com/company/{profile_id}"

    @staticmethod
    def build_people_search_url(
        company_name: Optional[str] = None,
        company_url: Optional[str] = None,
        keywords: Optional[str] = None,
    ) -> str:
        """
        Build a LinkedIn people-search URL for employees of a given company.

        Uses the company's LinkedIn URL as the currentCompany filter if available,
        otherwise falls back to a keyword search.
        """
        from urllib.parse import quote_plus

        # Use the company's /people/ page when a LinkedIn URL is available.
        # This is more reliable than the currentCompany search filter, which
        # requires a numeric company ID rather than a URL slug.
        if company_url:
            m = re.search(r"linkedin\.com/company/([a-zA-Z0-9\-_%]+)", company_url, re.IGNORECASE)
            if m:
                company_slug = m.group(1)
                base = f"https://www.linkedin.com/company/{company_slug}/people/"
                if keywords:
                    base += f"?keywords={quote_plus(keywords)}"
                return base

        # Fallback: keyword-based search
        q = company_name or ""
        if keywords:
            q = f"{keywords} {q}".strip()
        return f"https://www.linkedin.com/search/results/people/?keywords={quote_plus(q)}&network=%5B%22S%22%5D"
