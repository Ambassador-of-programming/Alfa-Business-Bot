import flet as ft
import json
from pc_vision.main_code import for_excel
import asyncio
import multiprocessing
import os


async def general(page: ft.Page):
    class Start:
        def __init__(self):
            self.error_text = ft.Text(visible=False, color='RED')

            self.radio_group = ft.RadioGroup(content=ft.Column([
                ft.Radio(value="BlueStacks App Player 3", label="BlueStacks App Player 3", label_style=ft.TextStyle(color="black")),
                ft.Radio(value="LDPlayer", label="LDPlayer", label_style=ft.TextStyle(color="black")),
                ])
            )
            self.f11 = ft.Checkbox(label="Полноэкранный режим (F11)", label_style=ft.TextStyle(color="black"))
            self.start_app = ft.CupertinoActivityIndicator(
                radius=10,
                color=ft.colors.RED,
                animating=True,
                visible=False
            )

            self.current_task = None
            self.button_stop = ft.ElevatedButton(text='Stop', on_click=self.event_button_stop, visible=False)
            self.button_start = ft.ElevatedButton(text='Start', on_click=self.event_button_start)
            self.reminder_text = ft.Text(value='Не забудьте использовать английскую раскладку клавиатуры и \nвключите эмулятор', color='GREEN')


        async def event_button_start(self, event):
            with open(file='config/conf.json', mode='r', encoding='utf-8') as file:
                data = json.load(file)

            if self.radio_group.value != None and data.get('filename_excel') != None and data.get('TesseractOCR') != None:
                self.start_app.visible = True
                self.button_start.visible = False
                self.button_stop.visible = True
                page.update()

                # Запускаем for_excel в отдельном процессе
                self.process = multiprocessing.Process(target=for_excel, args=(
                    data.get('filename_excel'),
                    self.radio_group.value,
                    self.f11.value,
                    data.get('TesseractOCR')
                ))
                self.process.start()

                # Ожидаем завершения процесса
                while self.process.is_alive():
                    await asyncio.sleep(0.1)
                    if not self.button_stop.visible:  # Проверяем, была ли нажата кнопка Stop
                        break

                if self.process.is_alive():
                    self.process.terminate()
                    self.process.join()

                self.start_app.visible = False
                self.button_start.visible = True
                self.button_stop.visible = False 
                page.update()

            else:
                self.error_text.visible = True
                self.error_text.value = "Выберите режим старта или выберите файл Excel или TesseractOCR"
                self.error_text.color = 'RED'
                page.update()

                await asyncio.sleep(3)
                self.error_text.visible = False
                page.update()
        
        async def event_button_stop(self, event):
            if self.process and self.process.is_alive():
                self.process.terminate()
                self.process.join()
            self.button_stop.visible = False
            self.button_start.visible = True
            self.start_app.visible = False
            page.update()

    start = Start()

    return ft.Column(controls=[
        ft.ResponsiveRow(controls=[
            start.radio_group,
            start.f11,
            start.error_text,
            start.start_app,
            start.button_start,
            start.button_stop,    
        ]),
        ft.Row(controls=[
            start.reminder_text,
        ], alignment=ft.MainAxisAlignment.CENTER)
    ])

# Предполагается, что у вас есть код для запуска Flet приложения, например:
# ft.app(target=general)