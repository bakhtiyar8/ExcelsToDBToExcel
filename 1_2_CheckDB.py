import sqlite3

# Подключаемся к базе данных SQLite
conn = sqlite3.connect('excel_data5.db')
c = conn.cursor()

# Проверяем, существует ли таблица excel_rf_old_data
c.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='excel_rf_old_data'")
if c.fetchone():
    print("Table excel_rf_old_data exists.")
    # Выполнение SQL-запроса для удаления пробелов в столбце "Antenna / Антенна"
    c.execute("UPDATE excel_rf_old_data SET [RF OLD DATA: Antenna / Антенна] = REPLACE([RF OLD DATA: Antenna / Антенна], ' ', '')")
    # Подтверждение изменений
    conn.commit()
    
    # Выполняем SQL-запрос для выборки данных из таблицы excel_data
    c.execute("SELECT * FROM excel_rf_old_data")
    # Получаем все строки с данными из таблицы
    rows = c.fetchall()
    
    if rows:
        print("Data in excel_rf_old_data:")
        # Выводим данные на экран
        for row in rows:
            print(row)
    else:
        print("No data found in excel_rf_old_data.")
else:
    print("Table excel_rf_old_data does not exist.")

# Проверяем, существует ли таблица RF_NEW_DATA
c.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='RF_NEW_DATA'")
if c.fetchone():
    print("Table RF_NEW_DATA exists.")
    # Выполнение SQL-запроса для удаления пробелов в столбце "Antenna / Антенна"
    c.execute("UPDATE RF_NEW_DATA SET [RF NEW DATA: Antenna / Антенна] = REPLACE([RF NEW DATA: Antenna / Антенна], ' ', '')")
    # Подтверждение изменений
    conn.commit()
    
    # Выполняем SQL-запрос для выборки данных из таблицы excel_data
    c.execute("SELECT * FROM RF_NEW_DATA")
    # Получаем все строки с данными из таблицы
    rows = c.fetchall()
    
    if rows:
        print("Data in RF_NEW_DATA:")
        # Выводим данные на экран
        for row in rows:
            print(row)
    else:
        print("No data found in RF_NEW_DATA.")
else:
    print("Table RF_NEW_DATA does not exist.")

# Проверяем, существует ли таблица Site_id_information
c.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='Site_id_information'")
if c.fetchone():
    print("Table Site_id_information exists.")
    # Выполняем SQL-запрос для выборки данных из таблицы excel_data
    c.execute("SELECT * FROM Site_id_information")
    # Получаем все строки с данными из таблицы
    rows = c.fetchall()
    
    if rows:
        print("Data in Site_id_information:")
        # Выводим данные на экран
        for row in rows:
            print(row)
    else:
        print("No data found in Site_id_information.")
else:
    print("Table Site_id_information does not exist.")

# Проверяем, существует ли таблица TSS_date
c.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='TSS_date'")
if c.fetchone():
    print("Table TSS_date exists.")
    # Выполняем SQL-запрос для выборки данных из таблицы excel_data
    c.execute("SELECT * FROM TSS_date")
    # Получаем все строки с данными из таблицы
    rows = c.fetchall()
    
    if rows:
        print("Data in TSS_date:")
        # Выводим данные на экран
        for row in rows:
            print(row)
    else:
        print("No data found in TSS_date.")
else:
    print("Table TSS_date does not exist.")

# Закрываем соединение с базой данных
conn.close()
