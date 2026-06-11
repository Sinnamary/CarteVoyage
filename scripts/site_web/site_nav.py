"""Fragments HTML partagés entre les pages du site."""

HEADER_NAV = """<header class="app-header">
    <a href="index.html" class="header-brand">
      <img src="assets/img/logo-cartevoyage.svg" alt="CarteVoyage" class="header-logo">
    </a>
    <nav class="header-nav" aria-label="Navigation principale">
      <a href="index.html" class="header-nav-link{active_map}">Carte</a>
      <a href="stats.html" class="header-nav-link{active_stats}">Statistiques</a>
      <a href="inspect.html" class="header-nav-link{active_inspect}">Contrôle</a>
    </nav>
  </header>"""


def render_header(active: str = "map") -> str:
    return HEADER_NAV.format(
        active_map=' is-active' if active == "map" else "",
        active_stats=' is-active' if active == "stats" else "",
        active_inspect=' is-active' if active == "inspect" else "",
    )
