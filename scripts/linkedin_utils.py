"""
LinkedIn research utilities.

All parsing functions are pure — they accept already-fetched page text (str)
and return structured data.  Browser navigation is handled externally (Claude
drives the browser via the Claude-in-Chrome MCP and passes text here).

Injectable fetch_fn:
    LinkedInResearcher(fetch_fn=my_fetcher)
    where my_fetcher(url: str) -> str returns page accessibility text.
    If fetch_fn is None, methods that need fetching raise BrowserNotWiredError.
"""

import difflib
import re
import time
from dataclasses import dataclass
from typing import Callable, Dict, List, Optional, Tuple

from url_parser import is_company_people_url, paginate_search_url

# Generic company-name tokens that carry no identifying signal — ignored when
# comparing a card's company against the target so e.g. "Navan Inc." still matches
# "Navan". Common suffixes ("group", "holdings", "labs", "technologies") are
# included deliberately: token overlap is only a *keep* signal, and dropping these
# stops the current-company filter from treating two unrelated firms that merely
# share a suffix ("Quantum Technologies" vs "Riverside Technologies") as a match.
# Substring and fuzzy matching still cover the legitimate same-company cases.
_COMPANY_STOPWORDS = {
    "the", "a", "an", "and", "inc", "ltd", "llc", "co", "corp", "company",
    "group", "gmbh", "plc", "holdings", "technologies", "labs",
}


def _company_matches(target: str, card_company: str) -> bool:
    """
    Conservative company-correctness test — biased hard toward *keeping* a card.

    Returns True (keep) when the card should not be dropped, False only on a
    *clear* company mismatch. The LinkedIn people tab leaks people whose current
    employer is not the target; this drops those while never false-dropping a
    legitimate hit.

    Keep when:
      * the target or the card company is missing/unparseable (we can't judge),
      * either name is a case-insensitive substring of the other
        ("Amazon" ⊂ "Amazon Web Services"),
      * the card company is a short all-caps abbreviation we can't resolve
        without an alias map (e.g. "AWS" under "Amazon") — bias to keep,
      * they share a significant (non-stopword) token,
      * their fuzzy ratio is ≥ 0.80.
    Drop only when none of the above hold.

    Subsidiary/abbreviation aliasing (AWS↔Amazon) is intentionally *not* resolved
    here — input-company correctness is the caller's responsibility.
    """
    if not target or not card_company:
        return True

    a = card_company.strip().lower()
    b = target.strip().lower()
    if a in b or b in a:
        return True

    # Short all-caps abbreviation we can't resolve — keep rather than false-drop.
    stripped = card_company.strip()
    if len(stripped) <= 4 and stripped.isupper():
        return True

    ta = {t for t in re.split(r"[^a-z0-9]+", a) if len(t) >= 3 and t not in _COMPANY_STOPWORDS}
    tb = {t for t in re.split(r"[^a-z0-9]+", b) if len(t) >= 3 and t not in _COMPANY_STOPWORDS}
    if ta & tb:
        return True

    return difflib.SequenceMatcher(None, a, b).ratio() >= 0.80


def _text_mentions_company(page_text: str, company: str) -> bool:
    """
    True when ``page_text`` names ``company`` — by a significant token, or the
    whole name when it has none. Returns True when no company is given, so the
    caller's guard is a no-op unless an expected company is supplied.

    This answers *only* "does the text name the company". Detecting a removed /
    placeholder page is a separate concern — see ``_looks_like_unavailable_page``.
    """
    if not company:
        return True
    tl = page_text.lower()
    tokens = [
        t for t in re.split(r"[^a-z0-9]+", company.lower())
        if len(t) >= 3 and t not in _COMPANY_STOPWORDS
    ]
    if not tokens:
        return company.lower() in tl
    return any(t in tl for t in tokens)


def _looks_like_unavailable_page(page_text: str) -> bool:
    """
    Positive signal that a company URL resolved to a removed / placeholder page
    (a ``/company/unavailable/`` redirect or a "page not found" shell) rather than
    merely a page that doesn't name the expected company. The org guard hard-skips
    these with a distinct, actionable message. Best-effort by design — the token
    check in the guard still rejects a page that never names the company/slug, so
    localized error wording that slips past these markers is caught there.
    """
    tl = page_text.lower()
    return (
        "/company/unavailable" in tl
        or "this page doesn" in tl
        or "page doesn't exist" in tl
        or "page not found" in tl
        or "page isn't available" in tl
    )


def _people_url_company_terms(search_url: str) -> str:
    """
    Return the company slug from a ``/company/<slug>/people/`` URL as
    space-separated terms ("amazon-web-services" → "amazon web services"), or ""
    if none can be extracted.

    LinkedIn slugs are derived from the company's real identity, so they match the
    page's *display* name even when the spreadsheet uses a different label (sheet
    "Alphabet" vs slug ``google`` → page "Google"). The org guard accepts a match
    on these terms too, so a correct URL with a differently-worded sheet name is
    not skipped.
    """
    m = re.search(r"/company/([a-zA-Z0-9\-_%]+)", search_url, re.IGNORECASE)
    if not m:
        return ""
    return m.group(1).replace("-", " ").replace("_", " ")


# Lines that mark the start of the employee/relationship content — everything
# above the first of these is the company's own header.
_HEADER_STOP_PREFIXES = (
    "search employees", "where they live", "where they studied",
    "what they do", "people you may know",
)


def _people_page_header(page_text: str) -> str:
    """
    The top slice of a people-tab text, where LinkedIn prints the company's *own*
    identity (the title line and the header: name, tagline,
    "<Industry> <City> <N> followers <N> employees").

    The org check matches against this header, **not** the whole body: card titles
    further down can name *other* companies (e.g. a wrong-org redirect whose cards
    mention a different "ACT Security"), which would otherwise let a mismatched page
    pass on a coincidental token. Stops at the first employee/facet section, capped
    for safety.
    """
    header: List[str] = []
    for line in page_text.splitlines():
        low = line.strip().lower()
        if any(low.startswith(p) for p in _HEADER_STOP_PREFIXES):
            break
        header.append(line)
        if len(header) >= 15:
            break
    return "\n".join(header)


class BrowserNotWiredError(RuntimeError):
    """Raised when browser automation is needed but no fetch_fn is injected."""
    pass


@dataclass
class LinkedInConnection:
    """A relevant person found via a mutual contact."""
    name: str
    title: str
    company: str
    mutual_contact: str

    def format_for_excel(self) -> str:
        # Avoid duplicating the company when it is already embedded in the title
        # e.g. "VP Product @ AWS" → don't append "at AWS" a second time
        if self.company and self.company.lower() in self.title.lower():
            return f"{self.name} ({self.title})"
        return f"{self.name} ({self.title} at {self.company})"


class RoleMatcher:
    """
    Matches job title strings against a list of relevant-role regex patterns.

    Patterns come entirely from the per-user config (see ``user_config.py``):
    plain titles are word-boundary-anchored by the config layer, and advanced
    users may add raw regexes. There is no hidden built-in list, so the config
    is the single source of truth for what counts as a relevant role.
    """

    def __init__(self, patterns: List[str]):
        self.patterns = patterns
        self._compiled: List[re.Pattern] = [
            re.compile(p, re.IGNORECASE) for p in patterns
        ]

    def matches(self, title: str) -> bool:
        if not title:
            return False
        for pattern in self._compiled:
            if pattern.search(title):
                return True
        return False


# ---------------------------------------------------------------------------
# Pure parsing functions
# ---------------------------------------------------------------------------

def extract_person_profile(page_text: str, url: str = "") -> Optional[Dict]:
    """
    Parse person-profile page text into structured data.

    Returns None if the profile is not found.
    """
    if not page_text:
        return None
    if "profile not found" in page_text.lower() or "page not found" in page_text.lower():
        return None

    lines = [l.strip() for l in page_text.splitlines() if l.strip()]
    name = lines[0] if lines else "Unknown"

    # LinkedIn profile headline is usually a short line near the top
    title = ""
    company = ""
    for i, line in enumerate(lines[:10]):
        # Skip the name line
        if i == 0:
            continue
        # Skip UI chrome
        if any(x in line.lower() for x in ["linkedin", "sign in", "connect", "follow", "message"]):
            continue
        if not title:
            title = line
        elif not company:
            company = line
            break

    return {"name": name, "title": title, "company": company, "url": url}


def extract_company_profile(page_text: str, url: str = "") -> Optional[Dict]:
    """
    Parse company-profile page text into structured data.

    Handles several LinkedIn layouts:
      1. About-page label pattern: a line "Industry" followed by the value on the next line.
      2. Header stats line with a city: "<Industry> <City>, <State> <N> followers ..."
      2b. Header stats line WITHOUT a city: "<Industry>  <N> followers  <N> employees".
      3. Inline "Industry: X" label.

    Industry defaults to "Unknown (verify)" — caller should show this to the user for confirmation.
    Returns None if profile not found.
    """
    if not page_text:
        return None
    if "profile not found" in page_text.lower() or "page not found" in page_text.lower():
        return None

    lines = [l.strip() for l in page_text.splitlines() if l.strip()]
    name = lines[0] if lines else "Unknown Company"

    industry = "Unknown (verify)"
    size = ""

    # Strategy 1: About-page two-line label format
    #   Industry            ← label line
    #   Software Development ← value line
    #   Company size        ← label line
    #   501-1,000 employees  ← value line
    for i, line in enumerate(lines):
        if line.lower() == "industry" and i + 1 < len(lines):
            candidate = lines[i + 1].strip()
            if candidate and not candidate.lower().startswith("company") and not candidate[0].isdigit():
                industry = candidate
        if line.lower() == "company size" and i + 1 < len(lines):
            size_candidate = lines[i + 1].strip()
            if size_candidate and not size:
                size = size_candidate

    # Strategy 2: Header stats line "<Industry> <City>, <State> <N> followers"
    # e.g. "Software Development Boston, Massachusetts 54K followers 501-1K employees"
    if industry == "Unknown (verify)":
        for line in lines:
            if "followers" in line.lower() and re.search(r"\d", line):
                # Extract everything before the first "City," pattern
                m = re.match(r"^(.+?)\s+[A-Z][a-z]+,", line)
                if m:
                    candidate = m.group(1).strip()
                    # Sanity-check: must look like an industry (no digits, reasonable length)
                    if candidate and not candidate[0].isdigit() and len(candidate) < 60:
                        industry = candidate
                        break

    # Strategy 2b: inline header stats WITHOUT a city, as LinkedIn renders the
    # company overview / people header:
    #   "Software Development  37M followers  10K+ employees"
    # Take the text before the follower count (e.g. "37M", "54K", "10K+", "501").
    if industry == "Unknown (verify)":
        for line in lines:
            if "followers" not in line.lower():
                continue
            m = re.match(r"^(.+?)\s+[\d.,]+[KMB]?\+?\s+followers", line, re.IGNORECASE)
            if m:
                candidate = m.group(1).strip()
                # Sanity-check: looks like an industry (no leading digit,
                # reasonable length, not itself a stats fragment).
                if (
                    candidate
                    and not candidate[0].isdigit()
                    and len(candidate) < 60
                    and "followers" not in candidate.lower()
                ):
                    industry = candidate
                    break

    # Strategy 3: inline "Industry: X" label
    if industry == "Unknown (verify)":
        for line in lines:
            m = re.match(r"(?:Industry|Sector)[:\s]+(.+)", line, re.IGNORECASE)
            if m:
                candidate = m.group(1).strip()
                if candidate:
                    industry = candidate
                    break

    # Size fallback: inline "Company size: X"
    if not size:
        for line in lines:
            m = re.match(r"(?:Company size|Employees)[:\s]+(.+)", line, re.IGNORECASE)
            if m:
                size = m.group(1).strip()
                break

    return {"name": name, "industry": industry, "size": size, "url": url}


# Module-level regexes for LinkedIn connection-degree markers.
# Two real layouts produce different markers (both validated against live data):
#   * /search/results/people/ cards use a bullet:        "• 2nd"
#   * /company/<slug>/people/ cards use a middle dot:     "· 2nd"
#     and additionally precede it with a spelled-out line: "2nd degree connection"
# Compiled once here to avoid re-compiling on every parse_search_results call.
_DEGREE_RE = re.compile(r"^[•·]\s*\d(?:st|nd|rd|th)\b", re.IGNORECASE)
_DEGREE_CONNECTION_RE = re.compile(
    r"^\d(?:st|nd|rd|th)\+?\s+degree\s+connection\b", re.IGNORECASE
)


def parse_search_results(page_text: str) -> List[Dict]:
    """
    Parse a LinkedIn people-search result page (from get_page_text) into cards.

    Real LinkedIn text format (validated against live data):
        Avery Lane
        • 2nd

        VP FP&A @ Coralogix
        Netanya, Center District, Israel
        Terry Wallace, Jordan Vance & 1 other mutual connection

        Victor Hale
        • 2nd
        ...

    Cards are separated by the degree marker "• 2nd" / "• 3rd" etc.
    The name line appears BEFORE the degree marker.
    Content (title, location, mutuals) follows the degree marker.
    """
    if not page_text:
        return []

    raw_lines = page_text.splitlines()
    results = []
    _UI_JUNK = {
        "linkedin", "sign in", "view profile", "show more results",
        "are these results helpful?", "next", "previous", "book an appointment",
        "people you may know", "connect", "follow", "message",
    }

    def is_junk(line: str) -> bool:
        stripped = line.strip().lower()
        return stripped in _UI_JUNK or stripped.startswith("page ") or not stripped

    for i, raw_line in enumerate(raw_lines):
        line = raw_line.strip()
        if not _DEGREE_RE.match(line):
            continue

        # Name is the nearest real line before the degree marker. Skip blanks, UI
        # junk, other degree markers, and the spelled-out "Nth degree connection"
        # line that the company /people/ layout inserts between name and marker.
        name = ""
        for j in range(i - 1, max(-1, i - 6), -1):
            candidate = raw_lines[j].strip()
            if (
                candidate
                and not is_junk(candidate)
                and not _DEGREE_RE.match(candidate)
                and not _DEGREE_CONNECTION_RE.match(candidate)
            ):
                name = candidate
                break
        if not name:
            continue

        # Collect content lines that follow, stopping at the next card — which
        # begins with either a degree marker ("· 2nd") or the spelled-out
        # "Nth degree connection" line (company /people/ layout).
        content: List[str] = []
        for j in range(i + 1, min(len(raw_lines), i + 12)):
            candidate = raw_lines[j].strip()
            if _DEGREE_RE.match(candidate) or _DEGREE_CONNECTION_RE.match(candidate):
                break
            if candidate and not is_junk(candidate):
                content.append(candidate)

        card = _parse_card_content(name, content)
        results.append(card)

    return results


def _parse_mutual_names(line: str) -> List[str]:
    """
    Extract all named mutual contacts from a mutual-connection line.

    Handles both real LinkedIn layouts (validated against live data):
      /search/results/ form (ampersand):
        "Jordan Vance is a mutual connection"
        "Jordan Vance & Drew Bennett are mutual connections"
        "Terry Wallace, Jordan Vance & 1 other mutual connection"
      /company/<slug>/people/ form (comma + "and", sometimes a followers prefix):
        "Devon Pierce is a mutual connection"
        "Wren Price, Marco Bianchi, and 1 other mutual connection"
        "16K followers • Martin Keller, Peter Zörner, and 3 other mutual connections"
    """
    # Strip a leading "<N> followers • " / "... followers · " stats prefix that the
    # company /people/ layout prepends to some cards' mutual line.
    line = re.sub(r"^.*?\bfollowers\b\s*[•·]\s*", "", line, flags=re.IGNORECASE)

    # Pattern 1: single mutual — "X is a mutual connection"
    m = re.match(r"^(.+?)\s+is a mutual connection", line, re.IGNORECASE)
    if m:
        return [m.group(1).strip()]

    # Pattern 2: named list + count — "Name1, Name2 & N other ..." or
    #            "Name1, Name2, and N other ..." (note: "&" or "and", optional comma)
    m = re.match(r"^(.+?),?\s*(?:&|and)\s+\d+\s+other\s+mutual", line, re.IGNORECASE)
    if m:
        names_str = m.group(1)
        return [n.strip() for n in re.split(r",\s*", names_str) if n.strip()]

    # Pattern 3: two or more named mutuals — "X & Y are mutual connections"
    #            or "X, Y are mutual connections" / "X and Y are mutual connections"
    m = re.match(r"^(.+?)\s+are\s+mutual\s+connections?", line, re.IGNORECASE)
    if m:
        names_str = m.group(1)
        parts = re.split(r",\s*|\s*&\s*|\s+and\s+", names_str)
        return [n.strip() for n in parts if n.strip()]

    return []


def _parse_card_content(name: str, content: List[str]) -> Dict:
    """Build a card dict from the name and its following content lines."""
    title = ""
    company = ""
    mutuals: List[str] = []

    for line in content:
        # Try mutual patterns first
        parsed_mutuals = _parse_mutual_names(line)
        if parsed_mutuals:
            mutuals.extend(parsed_mutuals)
            continue

        # Title (first non-location, non-mutual content line)
        if not title:
            title = line
            # Extract company from "Title @ Company" or "Title at Company" format
            m = re.search(r"\s+(?:@|at)\s+(.+)$", title, re.IGNORECASE)
            if m:
                company = m.group(1).strip()
            continue

        # Second content line is typically location — skip for our purposes

    return {"name": name, "title": title, "company": company, "mutuals": mutuals}


def find_connections_in_pages(
    page_texts: List[str],
    contact_name: str,
    role_matcher: RoleMatcher,
    default_company: str = "",
) -> List[LinkedInConnection]:
    """
    Given a list of already-fetched search-result page texts, return all people
    who have `contact_name` as a mutual contact and have a relevant role.

    contact_name may be a first name ("Toni") or full name ("Toni Vale-Hollis").
    Matching is case-insensitive substring on the mutual name field.
    default_company is used when the card does not include a company name (e.g. when
    the title is a plain string like "Senior Product Manager" without "@ Company").
    """
    connections: List[LinkedInConnection] = []
    seen: set = set()

    for page_text in page_texts:
        for card in parse_search_results(page_text):
            mutual_match = any(
                contact_name.lower() in m.lower() for m in card.get("mutuals", [])
            )
            if not mutual_match:
                continue
            if not role_matcher.matches(card.get("title", "")):
                continue

            # Drop cards clearly at a different company than the target — the
            # people tab leaks other-company people in. Conservative: keeps
            # unparseable / substring / abbreviation / fuzzy matches (see
            # _company_matches); only a clear mismatch is dropped. A missing
            # default_company disables the filter (unchanged behaviour).
            if not _company_matches(default_company, card.get("company", "")):
                continue

            key = (card["name"], card["title"])
            if key in seen:
                continue
            seen.add(key)

            company = card.get("company") or default_company
            connections.append(LinkedInConnection(
                name=card["name"],
                title=card["title"],
                company=company,
                mutual_contact=contact_name,
            ))

    return connections


# ---------------------------------------------------------------------------
# Researcher — thin orchestration wrapper
# ---------------------------------------------------------------------------

class LinkedInResearcher:
    """
    Thin wrapper that calls pure parsing functions with browser-fetched text.

    fetch_fn: callable(url: str) -> str  (must be injected for live use)
    """

    def __init__(self, fetch_fn: Optional[Callable[[str], str]] = None):
        self.fetch_fn = fetch_fn
        self.role_matcher: Optional[RoleMatcher] = None
        self.last_request_time: float = 0.0
        self.min_delay: float = 1.5  # seconds between browser requests

    def set_role_patterns(self, patterns: List[str]) -> None:
        self.role_matcher = RoleMatcher(patterns)

    # ------------------------------------------------------------------ #
    # Methods that need browser access                                     #
    # ------------------------------------------------------------------ #

    def _fetch(self, url: str) -> str:
        if self.fetch_fn is None:
            raise BrowserNotWiredError(
                "No fetch_fn provided — wire Claude-in-Chrome MCP before calling "
                "fetch methods, or supply page text directly to the pure parsers."
            )
        elapsed = time.time() - self.last_request_time
        if elapsed < self.min_delay:
            time.sleep(self.min_delay - elapsed)
        self.last_request_time = time.time()
        return self.fetch_fn(url)

    def fetch_person_profile(self, url: str) -> Optional[Dict]:
        text = self._fetch(url)
        return extract_person_profile(text, url)

    def fetch_company_profile(self, url: str) -> Optional[Dict]:
        text = self._fetch(url)
        return extract_company_profile(text, url)

    def search_company_employees(
        self,
        search_url: str,
        max_pages: int = 5,
        expected_company: Optional[str] = None,
    ) -> List[str]:
        """
        Fetch search-result page text for one company, returning raw page texts.

        Two LinkedIn surfaces need different handling (validated live 2026-07-04):

        * **Company people tab** (``/company/<slug>/people/``): mutual-connection
          cards load lazily in a "People you may know" block that is exhausted by
          scrolling and clicking **"Show more results"**, *not* by ``?page=N``.
          Claude does that expansion during FETCH and stores the single, fully
          expanded page text under this base URL, so here we fetch it exactly
          once. (The old ``len(text) < 200`` early-stop is gone — a freshly
          navigated page returns almost no cards and would abort immediately.)
        * **Keyword search fallback** (``/search/results/people/``): a genuine
          paginated surface, so we still walk ``?page=1..max_pages`` and stop when
          a page is empty or LinkedIn says it is showing all results.

        ``expected_company`` is a light org-correctness backstop to the FETCH-time
        title check: when supplied and the fetched people-tab text does not name
        that company (a wrong slug or a ``/company/unavailable/`` redirect), the
        page is skipped with a warning rather than silently parsed as empty.
        """
        if is_company_people_url(search_url):
            text = self._fetch(search_url)
            if not text or "no results" in text.lower():
                return []
            if expected_company:
                # Two distinct failure modes, each hard-skipped (keeping a wrong
                # page would attribute another company's people to the target —
                # cards with no parseable company would slip through the per-card
                # filter). The name/slug check runs against the page *header* only
                # (the company's own identity), not the whole body, so a redirect
                # to a different company whose cards happen to mention the target's
                # words is still caught. We accept the header when it names the
                # expected company OR the URL slug: the slug carries LinkedIn's own
                # name, so a correct URL with a differently-worded sheet label
                # ("Alphabet" vs slug "google" → page "Google") still matches.
                if _looks_like_unavailable_page(text):
                    print(
                        f"  ⚠ org check: '{expected_company}' resolved to an "
                        f"unavailable/removed page (check the LinkedIn URL) — "
                        f"skipping this company."
                    )
                    return []
                header = _people_page_header(text)
                slug_terms = _people_url_company_terms(search_url)
                if not (
                    _text_mentions_company(header, expected_company)
                    or (slug_terms and _text_mentions_company(header, slug_terms))
                ):
                    print(
                        f"  ⚠ org check: the people tab did not resolve to "
                        f"'{expected_company}' or its URL slug (header names a "
                        f"different company) — skipping to avoid mixing in another "
                        f"company's connections; verify the LinkedIn URL."
                    )
                    return []
            return [text]

        page_texts: List[str] = []
        for page_num in range(1, max_pages + 1):
            url = paginate_search_url(search_url, page_num)
            text = self._fetch(url)
            if not text or "no results" in text.lower():
                break
            page_texts.append(text)
            if "showing all results" in text.lower():
                break
        return page_texts

    def find_connections(
        self,
        company_name: str,
        contact_name: str,
        role_patterns: List[str],
        search_url: str,
        max_pages: int = 5,
    ) -> List[str]:
        """
        Full online search: fetch pages → parse → filter.
        Returns list of formatted Excel strings.
        """
        if self.role_matcher is None:
            self.set_role_patterns(role_patterns)

        page_texts = self.search_company_employees(
            search_url, max_pages, expected_company=company_name
        )
        connections = find_connections_in_pages(
            page_texts, contact_name, self.role_matcher, default_company=company_name
        )
        return [c.format_for_excel() for c in connections]
