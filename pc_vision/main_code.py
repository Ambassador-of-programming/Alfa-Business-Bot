import pytesseract
import cv2
import pyautogui
import time
from pc_vision.other_code import pin_code, process_excel, add_value_to_excel, activate_window
import multiprocessing


def check_screen(search_word, TesseractOCR: str):
    pytesseract.pytesseract.tesseract_cmd = rf"{TesseractOCR}"
    
    screenshot = pyautogui.screenshot(imageFilename='pc_vision/scrin.jpg')
    img = cv2.imread('pc_vision/scrin.jpg')
    img = cv2.resize(img, None, fx=3, fy=3, interpolation=cv2.INTER_CUBIC)

    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    gray = cv2.bilateralFilter(gray, 11, 17, 17)
    gray = cv2.Canny(gray, 30, 200)
    
    custom_config = r'--oem 3 --psm 6'
    data = pytesseract.image_to_data(gray, lang='rus+eng', 
        output_type=pytesseract.Output.DICT, config=custom_config)
    
    # Коэффициент масштабирования для корректировки координат
    scale_factor = 3
    
    print("Распознанный текст:", ' '.join([word for word in data['text'] if word != '']))
    
    for i, word in enumerate(data['text']):
        if int(data['conf'][i]) > 60:  # Проверка уверенности распознавания
            if 'руслан'.lower() in word.lower() or 'здравствуйте'.lower() in word.lower():
                return 'руслан'
            
            if search_word.lower() == 'карты':
                if search_word.lower() in word.lower() or 'e-mail'.lower() in word.lower() or 'счёта'.lower() in word.lower():
                    # Корректируем координаты
                    x = int(data['left'][i] / scale_factor)
                    y = int(data['top'][i] / scale_factor)
                    w = int(data['width'][i] / scale_factor)
                    h = int(data['height'][i] / scale_factor)

                    center_x = x + w // 2
                    center_y = y + h // 2

                    # Проверяем, находятся ли координаты в пределах экрана
                    screen_width, screen_height = pyautogui.size()
                    if 0 <= center_x < screen_width and 0 <= center_y < screen_height:
                        return (center_x, center_y)
                    # else:
                    #     print(f"Внимание: координаты ({center_x}, {center_y}) находятся за пределами экрана.")   
                    # 
            else:
                 if search_word.lower() in word.lower():
                    # Корректируем координаты
                    x = int(data['left'][i] / scale_factor)
                    y = int(data['top'][i] / scale_factor)
                    w = int(data['width'][i] / scale_factor)
                    h = int(data['height'][i] / scale_factor)

                    center_x = x + w // 2
                    center_y = y + h // 2

                    # Проверяем, находятся ли координаты в пределах экрана
                    screen_width, screen_height = pyautogui.size()
                    if 0 <= center_x < screen_width and 0 <= center_y < screen_height:
                        return (center_x, center_y)
    return None

def for_excel(file_path: str, window_title: str, f11_status: bool, TesseractOCR: str):
    file_path = rf'{file_path}'
    # window_title = "BlueStacks App Player 3"
    values_list = process_excel(file_path=file_path)

    activate_window(window_title=window_title)
    time.sleep(1)

    if f11_status == True:
        pyautogui.press('f11')
        time.sleep(0.3)
    
    for value in values_list:
        while True:
            result = check_screen('карты', TesseractOCR=TesseractOCR)
            if result == 'руслан':
                pin_code(TesseractOCR=TesseractOCR)
                time.sleep(1)
                continue

            elif result:
                center_x, center_y = result
                pyautogui.click(center_x, center_y)
                time.sleep(0.5)
                pyautogui.hotkey('ctrl', 'a')
                pyautogui.press('delete')
                pyautogui.typewrite(f"+7{value}")
                time.sleep(0.5)

                # Делаем новый скриншот после ввода номера
                pyautogui.screenshot(imageFilename='pc_vision/scrin.jpg')
                
                # Проверяем наличие слова "клиент" на новом скриншоте
                client_result = check_screen('клиент', TesseractOCR=TesseractOCR)
                if client_result:
                    print('есть клиент'.upper())
                    add_value_to_excel(file_path=file_path, search_value=int(value), value_to_add='Клиент Альфа-Банка')
                    # pyautogui.hotkey('ctrl', 'a')
                    # pyautogui.press('delete')
                    break  # Переходим к следующему значению в списке
                
                else:
                    print('нету клиент'.upper())
                    add_value_to_excel(file_path=file_path, search_value=int(value), value_to_add='Отсутствует')
                    # pyautogui.hotkey('ctrl', 'a')
                    # pyautogui.press('delete')
                    break  # Переходим к следующему значению в списке
            else:
                # continue
                time.sleep(0.5)  # Если ничего не найдено, ждем и проверяем снова

if __name__ == '__main__':
    # Этот блок нужен для корректной работы multiprocessing на Windows
    multiprocessing.freeze_support()


# Пример использования
# file_path = r'C:\Users\pro\Documents\project\Arkadii1988\Лист Microsoft Excel.xlsx'
# for_excel(file_path=file_path, window_title= "BlueStacks App Player 3")