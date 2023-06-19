"""Конфиг стилей интерфейса"""

LABELS_PROPERTIES = [  # Стили для лейблов интерфейса
    {"text": "Эта программа предназначена для выгрузки обменных форм из файлов project в папку",
     "relx": 0.17, "rely": 0.05},
    {"text": "Эта кнопка позволяет выбрать папку, в которую нужно выгрузить обменные формы",
     "relx": 0.2, "rely": 0.17},
    {"text": "Эта кнопка позволяет выбрать папку с файлами project для выгрузки обменной формы",
     "relx": 0.2, "rely": 0.27},
    {"text": "Эта кнопка позволяет начать выполнение программы",
     "relx": 0.2, "rely": 0.37},
    {"text": "Выгружено: 0 файлов",
     "relx": 1.2, "rely": 1.37},
    {"text": "Пожалуйста, ожидайте, выгрузка ОФ может занимать длительное время",
     "relx": 1.2, "rely": 1.37}
]

BUTTON_STYLE_ACTIVE = {'background': '#1166EE', 'foreground': 'white',  # Стиль активной кнопки
                       'font': ('Arial', 12)}
BUTTON_STYLE_DONE = {'background': '#118844', 'foreground': 'white',  # Стиль для успешно выполненной кнопки
                     'font': ('Arial', 12)}
BUTTON_STYLE_BLOCK = {'background': '#969699', 'foreground': 'white',  # Стиль для заблокированной кнопки
                      'font': ('Arial', 12)}