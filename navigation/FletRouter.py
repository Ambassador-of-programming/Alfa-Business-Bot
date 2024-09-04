import flet as ft
from pages.settings import settings
from pages.general import general

class Router:
    async def init(self, page: ft.Page):
        self.routes = {
            '/main': await general(page),
            '/settings': await settings(page),
        }
        self.body = ft.Container()

    async def route_change(self, route):
        self.body.content = self.routes[route.route]
        self.body.update()