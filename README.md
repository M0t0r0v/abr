13.11.2023 17:28
Not working yes, because not finished migration on words in DB

13.11.2023
Add Docstrings
Preparing DB for extend sql query
Preparing code fro extend sql query 


In the first version, the dictionary was in Excel, I converted dictionary to SQL and got better speed results.
Cleaned up my code by PEP8, now looks much better.
Think about how to optimize the code.

TO_DO
Изменить имена переменных сделать более понятными -
Имена функций сделать более понятным по смыслу
Больше докстрингов
Разбить базу слов по таблицам букв для оптимального поиска слов.
дел много времени мало...


избавился от дубликатов в словах (посмотрим, потестируем)
list_c = list(dict.fromkeys(list_c))

