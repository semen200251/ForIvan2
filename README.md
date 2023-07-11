# Программа для выгрузки обменных форм из файлов MS Project

## Описание программы
Программа выгружает обменные формы из файлов MS Project. Взаимодействие с интерфейсом происходит с помощью кнопок и пояснений к ним. Для выгрузки графиков пользователю требуется выбрать 2 папки: папка с файлами MS Project и папка, в которую программа поместит обменные формы. В интерфейсе присутствует цветовое сопровождение. Программа использует копии файлов, и создает резервную папку в Documents на ПК.

## Как начать работу

### Клонирование репозитория
<div style="background-color: #f2f2f2; padding: 10px;max-width: 20px;">
  <pre>
    <code>
      git clone https://github.com/semen200251/ForIvan2.git
    </code>
  </pre>
  <button onclick="copyToClipboard()"></button>
</div>

### Переход в папку репозитория
<div style="background-color: #f2f2f2; padding: 10px;max-width: 20px;">
  <pre>
    <code>
      cd ForIvan2
    </code>
  </pre>
  <button onclick="copyToClipboard()"></button>
</div>

### Установка зависимостей
<div style="background-color: #f2f2f2; padding: 10px;max-width: 20px;">
  <pre>
    <code>
      pip install -r requirements.txt
    </code>
  </pre>
  <button onclick="copyToClipboard()"></button>
</div>

### Запуск программы
<div style="background-color: #f2f2f2; padding: 10px;max-width: 20px;">
  <pre>
    <code>
      python main.py
    </code>
  </pre>
  <button onclick="copyToClipboard()"></button>
</div>

## Как работать с программой
- Нажать кнопку, которая выделена синим цветом, чтобы выбрать папку, в которую будут выгружены обменные формы;
- Нажать кнопку, которая после успешного выбора папки для выгрузки обменных форм выделится синим цветом, и выбрать папку с файлами MS Project;
- Нажать следующую кнопку, которая выделится синим цветом, чтобы начать выгрузку;
- Дождаться выгрузки обменных форм, это может занимать длительное время;
- Станут доступны 2 кнопки: для просмотра резервной папки и для просмотра папки с выгруженными обменными формами;
- Чтобы корректно выгрузить следюущую партию обменных форм рекомендуется перезагрузить приложение.

## Требования и ограничения
- Так как код реализован с помощью библиотеки PyWin32, то для работы требуется ОС Windows;
- У пользователя должен быть установлен MS Project для корректной работы;
- У пользователя должен быть доступ к папке Documents, там будет создана резеврная папка;
- У пользователя должен быть доступ к выбранным для работы папкам и файлам в них;
- У пользователя должен быть закрыт MS Project во время работы с программой.

<script>
function copyToClipboard() {
  var textToCopy = document.querySelector("code");
  var tempTextArea = document.createElement("textarea");
  tempTextArea.value = textToCopy.innerText;
  document.body.appendChild(tempTextArea);
  tempTextArea.select();
  document.execCommand("copy");
  document.body.removeChild(tempTextArea);
}
</script>
