def insert_zapis(lst: list):
    """Функция записи данных в Excel файл"""
    import pandas

    # Импортирование скрипта возвращения текущей директории
    from Maxim.main import script_dir

    result = pandas.DataFrame({"Data": [*lst]})
    try:
        writer = pandas.ExcelWriter(script_dir + "\\Insert.xlsx", engine='xlsxwriter')
        result.to_excel(writer, sheet_name='Страница 1')
        writer.close()
    except Exception as ex:
        print(ex)
