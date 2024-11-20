import pandas as pd
import matplotlib.pyplot as plt
import re


def parse_month_year(x):
    """
    Преобразует строку месяца и года в объект datetime.

    Args:
        x (str): Строка, содержащая название месяца и год в формате "Месяц Год" (например, "Май 2021").

    Returns:
        pd.Timestamp: Объект типа pandas Timestamp, представляющий первое число указанного месяца и года.
        pd.NaT: Если формат строки не соответствует ожидаемому, возвращает NaT (Not a Time).
    """
    month_year_dict = {
        "Январь": 1, "Февраль": 2, "Март": 3, "Апрель": 4, "Май": 5, "Июнь": 6,
        "Июль": 7, "Август": 8, "Сентябрь": 9, "Октябрь": 10, "Ноябрь": 11, "Декабрь": 12
    }
    if isinstance(x, str):
        parts = x.split()
        if len(parts) == 2:
            month_name, year = parts
            if month_name in month_year_dict:
                month = month_year_dict[month_name]
                try:
                    return pd.to_datetime(f"01 {month} {year}", format='%d %m %Y')
                except Exception as e:
                    print(f"Ошибка при преобразовании месяца: {x}, ошибка: {e}")
    return pd.NaT  # Если формат не распознан, возвращаем NaT


def load_and_prepare_data(file_path):
    """
    Загружает и подготавливает данные из Excel-файла, преобразуя суммы и даты, а также добавляя столбец для месяца и года.

    Args:
        file_path (str): Путь к Excel-файлу, который необходимо загрузить.

    Returns:
        pd.DataFrame: Подготовленный DataFrame, где преобразованы данные, добавлен столбец 'month_year' с датой месяца и года, и обработаны суммы и даты.

    Raises:
        FileNotFoundError: Если файл по указанному пути не найден.
        ValueError: Если возникла ошибка при загрузке данных или их обработке.
    """
    try:
        df = pd.read_excel(file_path)
        # Преобразование сумм денежных средств
        df['sum'] = df['sum'].replace({',': '.', ' ': ''}, regex=True)
        df['sum'] = pd.to_numeric(df['sum'], errors='coerce')
        # Переводим суммы в копейки
        df['sum'] = (df['sum'] * 100).round().astype('Int64')

        # Преобразование дат (если есть даты)
        df['receiving_date'] = pd.to_datetime(df['receiving_date'], format='%d.%m.%Y', errors='coerce')

        # Добавление столбца для месяца и года сделки
        df['month_year'] = None
        current_month_year = None
        # Регулярное выражение для поиска месяца и года
        month_year_pattern = r'([А-Яа-я]+)\s(\d{4})'

        for index, row in df.iterrows():
            status = row['status']

            if isinstance(status, str):
                # Пытаемся найти месяц и год в строке status
                match = re.search(month_year_pattern, status)

                if match:
                    month = match.group(1)  # Название месяца
                    year = match.group(2)  # Год
                    current_month_year = f'{month} {year}'  # Формируем строку "Месяц Год"

            if current_month_year:
                df.at[
                    index, 'month_year'] = current_month_year  # Записываем найденный месяц и год в добавленный столбец

        # Преобразуем 'month_year' в действительный объект даты для сортировки
        df['month_year'] = df['month_year'].apply(parse_month_year)
        return df
    except FileNotFoundError:
        raise FileNotFoundError(f"Файл {file_path} не найден. Проверьте путь.")
    except Exception as e:
        raise ValueError(f"Ошибка при загрузке данных: {e}")


def calculate_july_revenue(df):
    """
    Вычисляет общую выручку за июль 2021 года.

    Args:
        df (pd.DataFrame): Исходный DataFrame, содержащий данные о выручке, статусах и датах.

    Returns:
        float: Общая выручка за июль 2021 года в рублях.
    """
    # Фильтрация по июлю 2021
    july_revenue = df[
        (df['month_year'].dt.month == 7) &  # Месяц = 7 (Июль)
        (df['month_year'].dt.year == 2021) &  # Год = 2021
        (df['status'] == 'ОПЛАЧЕНО')
        ]
    total_july_revenue = july_revenue['sum'].sum() / 100  # Перевод обратно в рубли
    return total_july_revenue


def plot_revenue_trend(df):
    """
    Строит график изменения выручки по месяцам.

    Args:
        df (pd.DataFrame): Исходный DataFrame.
    """
    # Фильтрация по месяцам
    df_filtered = df[df['status'] == 'ОПЛАЧЕНО']
    monthly_revenue = df_filtered.groupby('month_year')['sum'].sum() / 100  # Перевод в рубли

    plt.figure(figsize=(12, 6))
    bars = monthly_revenue.plot(kind='bar', color='skyblue', edgecolor='black')

    plt.title('Изменение выручки компании', fontsize=16)
    plt.ylabel('Выручка (рубли)', fontsize=12)

    # Словарь для русских месяцев
    months_dict = {
        1: "Январь", 2: "Февраль", 3: "Март", 4: "Апрель", 5: "Май", 6: "Июнь",
        7: "Июль", 8: "Август", 9: "Сентябрь", 10: "Октябрь", 11: "Ноябрь", 12: "Декабрь"
    }
    month_year_labels = [
        f"{months_dict[date.month]} {date.year}" for date in monthly_revenue.index
    ]
    plt.xlabel('Месяц', fontsize=12)
    plt.xticks(ticks=range(len(monthly_revenue)), labels=month_year_labels, rotation=45, fontsize=10)
    plt.yticks(fontsize=10)
    plt.grid(axis='y', linestyle='--', alpha=0.7)

    for bar in bars.patches:
        value = bar.get_height()
        plt.text(bar.get_x() + bar.get_width() / 2, value + 5000,
                 f'{value:,.2f}', ha='center', fontsize=10, color='black')

    plt.tight_layout()
    plt.savefig('docs/revenue_trend.png', bbox_inches='tight')
    plt.show()


def find_top_manager(df):
    """
    Определяет лучшего менеджера за сентябрь 2021 года.

    Args:
        df (pd.DataFrame): Исходный DataFrame.

    Returns:
        tuple: Имя лучшего менеджера и его выручка.
    """
    # Фильтрация по сентябрю 2021
    september_revenue = df[
        (df['month_year'].dt.month == 9) &  # Месяц = Сентябрь
        (df['month_year'].dt.year == 2021) &  # Год = 2021
        (df['status'] == 'ОПЛАЧЕНО')  # Только оплаченные
        ]

    # Группировка по менеджеру и подсчет выручки
    manager_revenue = september_revenue.groupby('sale')['sum'].sum()

    # Если данных нет, возвращаем None и 0
    if manager_revenue.empty:
        return None, 0

    top_manager = manager_revenue.idxmax()
    top_revenue = manager_revenue.max() / 100  # Перевод суммы в рубли

    return top_manager, top_revenue


def find_dominant_deal_type(df):
    """
    Определяет преобладающий тип сделок в октябре 2021 года.

    Args:
        df (pd.DataFrame): Исходный DataFrame.

    Returns:
        str: Преобладающий тип сделок.
    """
    # Фильтрация по октябрю 2021
    october_deals = df[
        (df['month_year'].dt.month == 10) &  # Месяц = Октябрь
        (df['month_year'].dt.year == 2021)  # Год = 2021
        ]

    deal_type_counts = october_deals['new/current'].value_counts()

    # Возвращаем преобладающий тип сделки или None, если данных нет
    return deal_type_counts.idxmax() if not deal_type_counts.empty else None


def count_originals_in_june(df):
    """
    Считает количество оригиналов договоров по майским сделкам, полученных в июне 2021 года.

    Args:
        df (pd.DataFrame): Исходный DataFrame.

    Returns:
        int: Количество оригиналов.
    """
    # Фильтрация сделок за май 2021 года
    may_deals = df[
        (df['month_year'].dt.month == 5) &  # Месяц = Май
        (df['month_year'].dt.year == 2021)  # Год = 2021
        ]

    # Фильтрация оригиналов, полученных в июне 2021 года
    may_originals_june = may_deals[
        (may_deals['document'] == 'оригинал') &  # Тип документа = "оригинал"
        (may_deals['receiving_date'].dt.month == 6) &  # Месяц получения = Июнь
        (may_deals['receiving_date'].dt.year == 2021)  # Год получения = 2021
        ]

    # Возвращаем количество оригиналов
    return may_originals_june.shape[0]


def calculate_bonus(df):
    """
    Вычисляет бонусы для каждого менеджера на 01.07.2021 по описанным правилам.

    Args:
        df (pd.DataFrame): Исходный DataFrame с данными о сделках.

    Returns:
        pd.DataFrame: DataFrame с остатками бонусов на 01.07.2021 по менеджерам.
    """
    # Фильтрация сделок до июня 2021 года
    before_july_deals = df[
        (df['month_year'].dt.month < 7) &  # Месяц до июля (меньше 7)
        (df['month_year'].dt.year == 2021)  # Год = 2021
        ]

    # Фильтрация только оплаченных сделок с оригиналом договора для новых сделок
    new_deals = before_july_deals[(before_july_deals['new/current'] == 'новая') &
                                  (before_july_deals['status'] == 'ОПЛАЧЕНО') &
                                  (before_july_deals['document'] == 'оригинал') &
                                  (before_july_deals['receiving_date'].dt.month < 7)
        # Оригинал документа приходит до июля
                                  ].copy()


    # Бонусы для новых сделок
    new_deals.loc[:, 'bonus'] = new_deals.apply(lambda row: row['sum'] * 0.07 / 100, axis=1)

    # Для текущих сделок
    current_deals = before_july_deals[(before_july_deals['new/current'] == 'текущая') &
                                      (before_july_deals['status'] != 'ПРОСРОЧЕНО') &
                                      (before_july_deals['document'] == 'оригинал') &
                                      (before_july_deals['receiving_date'].dt.month < 7)
        # Оригинал документа приходит до июля
    ].copy()

    # Бонусы для текущих сделок
    current_deals.loc[:, 'bonus'] = current_deals.apply(
        lambda row: row['sum'] * 0.05 / 100 if row['sum'] > 10000 else row['sum'] * 0.03 / 100, axis=1)

    # Суммируем бонусы для каждого менеджера
    total_bonuses = pd.concat([new_deals, current_deals])
    manager_bonuses = total_bonuses.groupby('sale')['bonus'].sum()

    # TODO: Реализовать логику для случаев, когда оригинал документа приходит позже июня
    # # Для сделок, где оригинал приходит позже (позже июля)
    # remaining_bonuses = df[(df['receiving_date'].dt.month > 6) &
    #                        (df['receiving_date'].dt.year == 2021) &
    #                        (df['document'] == 'оригинал')].copy()
    #
    #
    # remaining_bonuses.loc[:, 'bonus'] = remaining_bonuses.apply(
    #     lambda row: row['sum'] * 0.07 / 100 if row['new/current'] == 'новая' and row['status'] == 'ОПЛАЧЕНО'
    #     else (row['sum'] * 0.05 / 100 if row['new/current'] == 'текущая' and row['sum'] > 10000
    #           else row['sum'] * 0.03 / 100), axis=1)
    #
    # # Суммируем остаточные бонусы
    # remaining_manager_bonuses = remaining_bonuses.groupby('sale')['bonus'].sum()
    #
    # # Добавляем остаточные бонусы
    # for manager, bonus in remaining_manager_bonuses.items():
    #     if manager in manager_bonuses:
    #         manager_bonuses[manager] += bonus
    #     else:
    #         manager_bonuses[manager] = bonus

    return manager_bonuses


def main():
    """
    Основная функция для загрузки, подготовки данных и выполнения анализа.

    Загружает данные из Excel-файла, затем выполняет несколько задач:
    1. Вычисляет общую выручку за июль 2021 года.
    2. Строит график изменения выручки по месяцам.
    3. Определяет лучшего менеджера за сентябрь 2021 года.
    4. Определяет преобладающий тип сделок в октябре 2021 года.
    5. Считает количество оригиналов договоров по майским сделкам, полученных в июне 2021 года.
    6. Вычисляет бонусы для каждого менеджера.

    В случае возникновения ошибок в процессе загрузки данных или анализа, выводится соответствующее сообщение об ошибке.

    Args:
        None

    Returns:
        None
    """
    file_path = "data.xlsx"  # Относительный путь до файла с данными для анализа

    try:
        df = load_and_prepare_data(file_path)
    except Exception as e:
        print(f"Ошибка при загрузке данных: {e}")
        return

    try:
        # Задача 1
        july_revenue = calculate_july_revenue(df)
        print(f"1. Общая выручка за июль 2021: {july_revenue:,.2f} руб.")

        # Задача 2
        print("2. Изменение выручки компании по месяцам:")
        plot_revenue_trend(df)

        # Задача 3
        top_manager, top_revenue = find_top_manager(df)
        if top_manager:
            print(f"3. Лучший менеджер за сентябрь 2021: {top_manager}, выручка: {top_revenue:,.2f} руб.")
        else:
            print("3. В сентябре 2021 года нет оплаченных сделок.")

        # Задача 4
        dominant_deal_type = find_dominant_deal_type(df)
        if dominant_deal_type:
            print(f"4. Преобладающий тип сделок в октябре 2021: {dominant_deal_type}")
        else:
            print("4. В октябре 2021 года нет сделок.")

        # Задача 5
        originals_count = count_originals_in_june(df)
        print(f"5. Оригиналов договора по майским сделкам, полученных в июне 2021: {originals_count}")

        # Задача 6: Вычисление бонусов
        manager_bonuses = calculate_bonus(df)
        manager_bonuses = manager_bonuses.round(2)

        if isinstance(manager_bonuses, pd.Series):
            manager_bonuses = manager_bonuses.reset_index()
            manager_bonuses.columns = ['sale', 'bonus (руб.)']

        print("6. Бонусы для менеджеров на 01.07.2021:")
        print(manager_bonuses.to_string(index=False, header=True))


    except Exception as e:
        print(f"Ошибка в процессе анализа: {e}")


if __name__ == "__main__":
    main()
