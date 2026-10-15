import os

import reflex as rx

config = rx.Config(
    app_name="reflex_grid",
    plugins=[
        # Always light (as CoDoc in MD): the app paints white sheets, ribbon and dialogs, so the theme's
        # dark colours (light text) would vanish on them when the OS is in dark mode.
        rx.plugins.RadixThemesPlugin(
            theme=rx.theme(appearance="light", accent_color="cyan", gray_color="slate", radius="small", scaling="95%"),
        ),
        rx.plugins.TailwindV3Plugin(),
    ],
    disable_plugins=[rx.plugins.SitemapPlugin],
    # Behind a reverse proxy (e.g. re_ddns nginx) API_URL points the browser's websocket at it.
    api_url=os.getenv("API_URL", "http://localhost:8000"),
    cors_allowed_origins=["*"],
)
