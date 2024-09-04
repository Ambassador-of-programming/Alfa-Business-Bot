import flet as ft
import asyncio 
import json


async def settings(page: ft.Page):
    class Excel:
        def __init__(self):
            self.result_text_folder = ft.Text(visible=False)
            self.result_text_file = ft.Text(visible=False)

            self.text_save_folder = ft.Text(color='black', value='''Выберите файл с Tesseract-OCR. \nДолжно быть: Tesseract-OCR/tesseract.exe''')
            self.folder_picker = ft.FilePicker(on_result=self.on_pick_folder)
            self.folder_select = ft.ElevatedButton(
                text="Выбрать файл .exe", 
                on_click=lambda _: self.folder_picker.pick_files(
                    allow_multiple=True
                )
            )

            self.text_select_file_excel = ft.Text(color='black', value='''Выберите файл с данными Excel''')
            self.file_picker = ft.FilePicker(on_result=self.on_pick_file)
            self.file_excel_select = ft.ElevatedButton(
                "Выбрать Excel",
                    icon=ft.icons.UPLOAD_FILE,
                    on_click=lambda _: self.file_picker.pick_files(
                        allow_multiple=True
                    ),
            )

        async def on_pick_folder(self, event: ft.FilePickerResultEvent):
            if event.files != None:
                with open('config/conf.json', 'r+', encoding='utf-8') as file:
                    data = json.load(file)
                    data['TesseractOCR'] = event.files[0].path  # Обновляем значение

                    # Перематываем файл на начало и очищаем его перед записью
                    file.seek(0)
                    file.truncate()

                    json.dump(data, file, ensure_ascii=False, indent=4)

                self.result_text_folder.visible = True
                self.result_text_folder.value = f"Вы выбрали: {event.files[0].path}"
                self.result_text_folder.color = 'green'
                page.update()

                await asyncio.sleep(3)
                self.result_text_folder.visible = False
                page.update()
            else:
                self.result_text_folder.visible = True
                self.result_text_folder.value = "Папка не выбрана"
                self.result_text_folder.color ='red'
                page.update()

                await asyncio.sleep(3)
                self.result_text_folder.visible = False
                page.update()

        async def on_pick_file(self, event: ft.FilePickerResultEvent):
            if event.files != None:
                with open('config/conf.json', 'r+', encoding='utf-8') as file:
                    data = json.load(file)
                    data['filename_excel'] = event.files[0].path  # Обновляем значение

                    # Перематываем файл на начало и очищаем его перед записью
                    file.seek(0)
                    file.truncate()

                    json.dump(data, file, ensure_ascii=False, indent=4)

                self.result_text_file.visible = True
                self.result_text_file.value = f"Вы выбрали: {event.files[0].path}"
                self.result_text_file.color = 'green'
                page.update()

                await asyncio.sleep(3)
                self.result_text_file.visible = False
                page.update()

            else:
                self.result_text_file.visible = True
                self.result_text_file.value = "Файл не выбран"
                self.result_text_file.color ='red'
                page.update()

                await asyncio.sleep(3)
                self.result_text_file.visible = False
                page.update()

    excel = Excel()

    return ft.Column(controls=[
        excel.folder_picker,
        excel.file_picker,

        ft.ResponsiveRow(controls=[
            excel.text_save_folder,
            excel.result_text_folder,
            excel.folder_select,
        ]),

        ft.ResponsiveRow([ft.Divider(height=9, thickness=3),]),

        ft.ResponsiveRow(controls=[
            excel.text_select_file_excel,
            excel.result_text_file,
            excel.file_excel_select,
        ]),

        ft.ResponsiveRow([ft.Divider(height=9, thickness=3),]),
        
    ])