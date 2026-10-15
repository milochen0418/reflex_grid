import os

import reflex as rx

config = rx.Config(
    app_name="reflex_grid",
    plugins=[rx.plugins.TailwindV3Plugin()],
    disable_plugins=["reflex.plugins.sitemap.SitemapPlugin"],
    # Behind a reverse proxy (e.g. re_ddns nginx) API_URL points the browser's websocket at it.
    api_url=os.getenv("API_URL", "http://localhost:8000"),
    cors_allowed_origins=["*"],
)
