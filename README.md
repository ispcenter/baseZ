"заявочная1.0.ру" - программа для управления заявками.

Функции:

- занесение информации из заявок в базу данных (далее БД);
- генерация файла заявки из записей в базе, относящихся к указанной заявке.

Как пользоваться

1) Желательно, чтобы программа хранилась в отдельной папке, в которой должен также находиться файл
с именем "TTlogo.png" c логотипом "Турботехника", иначе его придется вставлять после генерации заявки вручную.
2) Как обычно двойным щелчком ЛКМ запустить файл "заявочная1.1.py", следовать советам программы.

Описание работы программы

При запуске функции дополнения БД программа проверяет наличие файла БД (bdz.xlsx) в свой папке:
- если его не окажется в момент загрузки заявки, программа создаст новую БД в своей папке;
- если файл БД имеется, то программа дополнит его новой заявкой.
Номер формата "N-YYYY" будет присвоен заявке автоматически путем определения года её оформления
и нахождением в базе последнего номера в этом году, например программа увидела дату оформления
заявки 13.05.2019, нашла в базе среди заявок 2019го года последний номер был "3-2019", присвоила
новой заявке номер "4-2019".

При генерации заявки её файл появится в папке программы. Название будет содержать номер заявки,
например "заявка_№1-2019.xlsx".
