"""Модуль отвечает за интерфейс приложения. Также в нем содержится управляющая функция всего проекта."""

import os
import time

import tkinter as tk
from tkinter import messagebox, filedialog, scrolledtext

import core.readOF as readOF
import settings.interface as config_for_interface
import settings.interface_style as interface_style
import core.io as oi


def choose_folder(folder_id, buttons):
    """Функция обрабатывает нажатие кнопки для различных ситуаций.

    На вход поступает идентификатор кнопки:
    1 - Кнопка для выбора папки, в которую будут выгружены обменные формы
    2 - Кнопка для выбора папки, в которой находятся файлы project
    Следующие кнопки находятся в другой вкладке приложения
    3 - Кнопка для выбора папки, в которой находятся обменные формы для внесения факта в файлы project
    4 - Кнопка для выбора папки, в которой находятся файлы project, ожидающие внесения факта.
    """
    folder_path = filedialog.askdirectory()
    if folder_id == 1:
        if folder_path:
            config_for_interface.path_to_to_folder = folder_path
            buttons[0].configure(bg="#118844")
            buttons[1].configure(state="normal", bg="#1166EE")
        else:
            messagebox.showerror("Ошибка", "Выберите папку")
            return
    elif folder_id == 2:
        if folder_path:
            config_for_interface.path_to_from_folder = folder_path
            buttons[1].configure(bg="#118844")
            buttons[2].configure(state="normal", bg="#1166EE")
        else:
            messagebox.showerror("Ошибка", "Выберите папку")
            return


def _get_paths_to_file(directory):
    """Функция для получения абсолютных путей до файлов в папке.

    На вход поступает путь до папки. Функция возвращает список с
    абсолютными путями до всех файлов лежащих в этой папке
    и во вложенных папках.
    """
    file_paths = []
    for root, directories, files in os.walk(directory):
        for file in files:
            file_path = os.path.abspath(os.path.join(root, file))
            file_paths.append(file_path)
    return file_paths


def _update_progress(value, count, labels):
    """Функция обновляет значение количества загруженных файлов.

    На вход поступает количество выгруженных обменных форм и количество,
    которое требуется выгрузить. Она выводит информацию на экран о
    статусе загрузки, сколько файлов из общего числа выгружено на
    данный момент.
    """
    labels[-2].configure(text=f"Выгружено: {value} файлов из {count}")


def _switch_info_labels(value, labels):
    """Функция выводит на экран текстовую информацию о результате загрузки.

    На вход поступает число, если его значение равно 0, то она выводит на
    экран информацию для процесса выгрузки. Если значение не равно 0,
    то на экран выводится информация о результате загрузки.
    """
    succes = sum(1 for item in config_for_interface.path_to_results
                 if item is not None)
    if value == 0:
        labels[-1].configure(text="Пожалуйста, ожидайте, выгрузка ОФ может занимать длительное время")
    else:
        labels[-1].configure(
            text=f"Загрузка завершена. Загружено: {len(config_for_interface.path_to_results)} файлов.\n Успешно: {succes}")


def _change_after_work(value, labels, buttons):
    """Функция задает изменения в стилях после окончания выгрузки обменных форм.

    На вход поступает целочисленное значение, с помощью него она запускает
    функцию для вывода информации о результате выгрузки обменных форм. Также
    она скрывает лейбл с информацией для процесса загрузки.
    """
    labels[-2].place_forget()
    _switch_info_labels(value, labels)
    labels[-1].place_configure(relx=0.025, rely=0.5)
    buttons[2].configure(bg="#118844")
    buttons[3].configure(state="normal", bg="#1166EE")
    buttons[4].configure(state="normal", bg="#1166EE")


def _find_name(list_projects, excel_path):
    """Функция ищет файл project, которому соответствует конкретная обменная форма.

    На вход поступает список с абсолютными путями до файлов project,
    и путь до обменной формы. На выход поступает абсолютный путь до файла project
    c таким же именем.
    """
    name = os.path.splitext(os.path.basename(excel_path))[0]
    for path in list_projects:
        file_name = os.path.splitext(os.path.basename(path))[0]
        if file_name == name:
            return path


def start_click(folder_id, labels, window, text_area, buttons):
    """Функция выполняет основной функционал

    На вход поступает id кнопки.
    folder_id = 1 - выполняется выгрузка обменных форм
    folder_id = 2 - выполняется внесение факта в файл project
    Перед выполнением основного функционала, функция выполняет подготовку.
    То есть создает резервную папку, получает абсолютные пути до нужных файлов.
    Затем вызывает управляющую функцию из модуля, в котором содержится нужный
    функционал. Затем сохраняет полученные результаты в необходимые папки.
    """
    if not os.path.exists(config_for_interface.PATH_TO_RESERVE_FOLDER):
        os.mkdir(config_for_interface.PATH_TO_RESERVE_FOLDER)
    path_to_excel_folder = config_for_interface.PATH_TO_RESERVE_FOLDER + '\\' + "OF"
    path_to_project_folder = config_for_interface.PATH_TO_RESERVE_FOLDER + '\\' + "projects"
    path_to_unsuccessful_folder = config_for_interface.PATH_TO_RESERVE_FOLDER + '\\' + "unsuccessful"
    paths_to_bad_files = []
    if not os.path.exists(path_to_excel_folder):
        os.mkdir(path_to_excel_folder)
    if not os.path.exists(path_to_project_folder):
        os.mkdir(path_to_project_folder)
    if not os.path.exists(path_to_unsuccessful_folder):
        os.mkdir(path_to_unsuccessful_folder)
    value = 0
    labels[-2].place(relx=0.025, rely=0.5)
    labels[-1].place(relx=0.025, rely=0.55)
    _switch_info_labels(value, labels)
    if folder_id == 1:
        paths_to_projects = _get_paths_to_file(config_for_interface.path_to_from_folder)
        if paths_to_projects is None:
            return
        _update_progress(value, len(paths_to_projects), labels)
        window.update()
        try:
            oi.transfer_files(paths_to_projects, path_to_project_folder)
        except Exception as e:
            print(e)
            messagebox.showerror("Ошибка",
                                 e)
            return
        for path in paths_to_projects:
            start = time.time()
            file_name = os.path.basename(path)
            path = os.path.join(path_to_project_folder, file_name)
            res = readOF.main(path, path_to_excel_folder)
            end = time.time()
            if res is None:
                paths_to_bad_files.append(path)
                text_area.insert(tk.INSERT,
                                 f"{os.path.basename(paths_to_projects[value])}    -    Не успешно\n")
            else:
                text_area.insert(tk.INSERT,
                                 f"{os.path.basename(paths_to_projects[value])}    -    Успешно"
                                 f"    Время: {int(end-start)} cекунд\n")
            value = value + 1
            _update_progress(value, len(paths_to_projects), labels)
            window.update()
            config_for_interface.path_to_results.append(res)
    try:
        oi.transfer_files(config_for_interface.path_to_results, config_for_interface.path_to_to_folder)
    except Exception as e:
        print(e)
        messagebox.showerror("Ошибка",
                             e)
        return
    if paths_to_bad_files:
        oi.transfer_files(paths_to_bad_files, path_to_unsuccessful_folder)

    _change_after_work(value, labels, buttons)


def on_window_resize(window, buttons, labels, text_area):
    """Обработчик события изменения размеров окна."""

    new_width = window.winfo_width()
    new_height = window.winfo_height()

    button_width = int(new_width / 7)
    button_height = int(new_height / 15)
    label_width = int(new_width / 1.2)
    label_height = int(new_height / 15)
    text_area_height = int(new_height / 3)
    buttons[0].place(width=button_width, height=button_height)
    buttons[1].place(width=button_width, height=button_height)
    buttons[2].place(width=button_width, height=button_height)
    labels[1].place(width=label_width, height=label_height)
    labels[2].place(width=label_width, height=label_height)
    labels[3].place(width=label_width, height=label_height)
    text_area.place(width=label_width / 1.3, height=text_area_height)


def open_reserve_folder():
    """Открывает резервную папку."""

    if config_for_interface.PATH_TO_RESERVE_FOLDER:
        os.startfile(config_for_interface.PATH_TO_RESERVE_FOLDER)


def open_folder_with_res():
    """Открывает папку с результатами работы."""

    if config_for_interface.path_to_to_folder:
        os.startfile(config_for_interface.path_to_to_folder)


def create_button(self, props):
    """Создает кнопку по переданным атрибутам кнопки, размещает ее и возвращает объект кнопки."""

    button = tk.Button(self, text=props["text"], command=props["command"], **props["style"], state=props["state"],
                           width=props["width"])
    button.place(relx=props["relx"], rely=props["rely"])
    return button


def create_label(self, props):
    """Создает лейбл по переданным атрибутам лейбла, размещает его и возвращает объект лейбла."""

    label = tk.Label(self, text=props["text"], font=('Arial', 12), background="light blue")
    label.place(relx=props["relx"], rely=props["rely"])
    return label


def configure_all(window, buttons, button_properties, labels):
    """Задает стили для окна приложения, кнопок и лейблов.

    На вход принимает экземпляр объекта приложения, список кнопок,
    список стилей кнопок, список лейблов
    На выход поступает список кнопок, список лейблов и текстовое поле для
    вывода лога приложения.
    """
    window.title("Приложение для работы с ОФ")
    window.geometry("1000x600")
    window.minsize(1000, 600)
    window.configure(background="light blue")

    for props in button_properties:
        buttons.append(create_button(window, props))
    for props in interface_style.LABELS_PROPERTIES:
        labels.append(create_label(window, props))
    text_area = scrolledtext.ScrolledText(window, width=80, height=10)
    text_area.place(relx=0.17, rely=0.6)
    return buttons, labels, text_area
