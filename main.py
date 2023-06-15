"""Управляющий модуль"""

from interface.interface import *

if __name__ == '__main__':
    window = tk.Tk()
    BUTTON_PROPERTIES = [
        {"text": "Выбрать папку", "command": lambda: choose_folder(1, buttons),
         "style": interface_style.BUTTON_STYLE_ACTIVE,
         "width": 15, "relx": 0.025, "state": "normal", "rely": 0.17},
        {"text": "Выбрать папку", "command": lambda: choose_folder(2, buttons),
         "style": interface_style.BUTTON_STYLE_BLOCK,
         "width": 15, "relx": 0.025, "state": "normal", "rely": 0.27},
        {"text": "Начать", "command": lambda: start_click(1, labels, window, text_area, buttons),
         "style": interface_style.BUTTON_STYLE_BLOCK,
         "width": 15, "relx": 0.025, "state": "normal", "rely": 0.37},
        {"text": "Открыть резервную папку", "command": open_reserve_folder,
         "style": interface_style.BUTTON_STYLE_BLOCK,
         "width": 21, "relx": 0.78, "state": "normal", "rely": 0.5},
        {"text": "Открыть папку с ОФ", "command": open_folder_with_res,
         "style": interface_style.BUTTON_STYLE_BLOCK,
         "width": 17, "relx": 0.6, "state": "normal", "rely": 0.5}
    ]
    buttons = list()
    labels = list()

    buttons, labels, text_area = configure_all(window, buttons, BUTTON_PROPERTIES, labels)

    messagebox.showwarning("Предупреждение",
                           "Пожалуйста, закройте открытые файлы project для корректной работы программы")
    window.mainloop()
