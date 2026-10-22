import pandas as pd

# 1. Читаем загруженный Excel файл
df = pd.read_excel('tariffs_template.xlsx')

# 2. Преобразуем из широкого формата (матрицы) в плоский (под базу данных)
# id_vars=['Тип ТС'] означает, что этот столбец остается, а зоны сворачиваются
df_melted = df.melt(
    id_vars=['Тип ТС'], 
    var_name='zone', 
    value_name='base_tariff'
)

# 3. Фильтруем пустые значения (если для какой-то зоны нет цены)
df_cleaned = df_melted.dropna(subset=['base_tariff'])

# 4. Переименовываем колонку ТС для стандарта БД
df_cleaned = df_cleaned.rename(columns={'Тип ТС': 'vehicle_type_code'})

# Результат готов для вставки в базу. 
# Пример строки: vehicle_type_code="B1", zone="1", base_tariff=1500.00