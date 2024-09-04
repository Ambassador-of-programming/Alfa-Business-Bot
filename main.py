import flet as ft
from navigation.FletRouter import Router
from navigation.bar import Appbar


async def main(page: ft.Page):
    page.title = 'Alfa-Business-Bot'
    page.theme_mode = ft.ThemeMode.DARK
    page.scroll = 'HIDDEN'
    page.padding = 10
    page.platform = ft.PagePlatform.WINDOWS
    page.window.width = 500
    page.window.height = 500
    
    page.adaptive = True

    myRouter = Router()
    await myRouter.init(page)

    appbar = Appbar(page)
    page.bottom_appbar = await appbar.content()

    page.bgcolor = '#F6F6F6'
    page.on_route_change = myRouter.route_change
    page.add(
        myRouter.body
    )
    
    page.go('/main')

if __name__ == "__main__":
    ft.app(main)