# Работа с Excel

Данный проект показывает, как можно считывать и записывать данные в формате Excel. Подробности вы можете прочитать в статьях [Чтение данных из Excel](https://devmark.ru/article/excel-reading) и [Сохранение данных в формате Excel](https://devmark.ru/article/excel-writing).

При запуске из консоли следует указать два параметра: действие (**read** или **write**) и абсолютный путь до файла.

Ещё больше статей по разработке ПО вы можете найти на моём сайте [devmark.ru](https://devmark.ru/).

## Примеры:
Чтение данных из файла:

```java -jar target/excel-example-1.0.jar read /home/user/file_for_read.xlsx```

Запись данных в файл:

```java -jar target/excel-example-1.0.jar write /home/user/file_for_write.xlsx```