import pandas as pd
import numpy as np
from datetime import timedelta
import warnings
warnings.filterwarnings('ignore') # Чтобы не засорять консоль предупреждениями pandas

# Настройки pandas для красивого вывода в консоль
pd.set_option('display.max_columns', None)
pd.set_option('display.width', 1000)

def print_header(title):
    print(f"\n{'='*80}\n{title}\n{'='*80}")
# ==============================================================================
# БЛОК 1: ЗАГРУЗКА И ПОДГОТОВКА ДАННЫХ
# Что делаем: Читаем все листы Excel, приводим даты к правильному формату.
# ==============================================================================
print_header("БЛОК 1: ЗАГРУЗКА И ПОДГОТОВКА ДАННЫХ")
print("Что делаем: Читаем Excel-файлы, проверяем структуру таблиц и приводим даты к единому формату.")

file_path = "candidate_data.xlsx"
sku_df = pd.read_excel(file_path, sheet_name="SKU")
hist_df = pd.read_excel(file_path, sheet_name="DailyHistory")
supply_df = pd.read_excel(file_path, sheet_name="OpenSupply")
calendar_df = pd.read_excel(file_path, sheet_name="OrderCalendar")

# Приводим колонки с датами к типу datetime
hist_df['Date'] = pd.to_datetime(hist_df['Date'])
supply_df['ExpectedReceiptDate'] = pd.to_datetime(supply_df['ExpectedReceiptDate'])
calendar_df['OrderDate'] = pd.to_datetime(calendar_df['OrderDate'])

STOCK_COL = "Остаток на конец дня"
AS_OF_DATE = pd.to_datetime("2025-09-19")
LEAD_TIME = 10 # Учебный срок поставки из условий

print("\n[INFO] Шапка SKU (Текущие остатки):")
print(sku_df.head(3))
print("\n[INFO] Шапка DailyHistory (История продаж):")
print(hist_df.head(3))
print("\n[INFO] Шапка OpenSupply (Открытые поставки):")
print(supply_df.head(3))
print("\n[INFO] Шапка OrderCalendar (Календарь заказов):")
print(calendar_df.head(3))

# ==============================================================================
# БЛОК 1.2: EDA (РАЗВЕДОЧНЫЙ АНАЛИЗ ДАННЫХ)
# ==============================================================================
print("\n" + "="*60)
print("EDA: ИССЛЕДОВАНИЕ ДАННЫХ")
print("="*60)

# 1. Общая информация
print(f"1. Общая статистика:")
print(f"   - Всего SKU в активном ассортименте: {sku_df['SKU'].nunique()}")
print(f"   - Период истории: с {hist_df['Date'].min().date()} по {hist_df['Date'].max().date()}")
print(f"   - Всего дней в истории: {hist_df['Date'].nunique()}")
print(f"   - Ожидаемое количество строк (SKU * Дни): {sku_df['SKU'].nunique() * hist_df['Date'].nunique()}")
print(f"   - Фактическое количество строк: {len(hist_df)}")

# 2. Проверка на пропуски (Missing Values)
print("\n2. Проверка на пропуски (NaN) в ключевых колонках:")
missing_sales = hist_df['SalesQty'].isna().sum()
missing_stock = hist_df[STOCK_COL].isna().sum()
print(f"   - Пропуски в SalesQty: {missing_sales} (шт.)")
print(f"   - Пропуски в {STOCK_COL}: {missing_stock} (шт.)")
if missing_stock > 0:
    print("     [ВНИМАНИЕ] Пустой остаток означает отсутствие наблюдения (см. Data Dictionary).")
    print("     В модели мы будем игнорировать такие дни при расчете ADS.")

# 3. Проверка на аномалии (Отрицательные остатки или продажи)
neg_sales = hist_df[hist_df['SalesQty'] < 0]
neg_stock = hist_df[hist_df[STOCK_COL] < 0]
print(f"\n3. Проверка на аномалии:")
print(f"   - Дней с отрицательными продажами: {len(neg_sales)}")
print(f"   - Дней с отрицательным остатком: {len(neg_stock)}")
if len(neg_stock) > 0:
    print("     [ВНИМАНИЕ] Отрицательный остаток может означать ошибку в данных или backorder.")

# 4. Топ-5 и Аутсайдеры по продажам
total_sales = hist_df.groupby('SKU')['SalesQty'].sum().sort_values(ascending=False)
print("\n4. Разброс спроса (Топ-5 и Худшие-5 по сумме продаж за весь период):")
print("   --- ТОП-5 (Локомотивы) ---")
for sku, val in total_sales.head(5).items():
    print(f"     {sku}: {val:.0f} шт.")
print("   --- ХУДШИЕ-5 (Мертвый груз) ---")
for sku, val in total_sales.tail(5).items():
    print(f"     {sku}: {val:.0f} шт.")

# 5. Анализ "Мертвых запасов" (Dead Stock)
# Ищем SKU, у которых есть текущий остаток на 19.09, но не было продаж за ВСЮ историю
print(f"\n5. Анализ 'Мертвых запасов' (на 19.09.2025):")

current_stock_df = sku_df[['SKU', 'CurrentStock']].copy()
dead_stock_check = current_stock_df.merge(total_sales.reset_index(), on='SKU', how='left')
dead_stock_check['SalesQty'] = dead_stock_check['SalesQty'].fillna(0)
dead_stock = dead_stock_check[(dead_stock_check['CurrentStock'] > 0) & (dead_stock_check['SalesQty'] == 0)]

print(f"   - SKU с остатком > 0, но без продаж за всю историю: {len(dead_stock)}")
if len(dead_stock) > 0:
    print(f"     Суммарный замороженный капитал (штук): {dead_stock['CurrentStock'].sum():.0f}")
    print(f"     Примеры SKU: {', '.join(dead_stock['SKU'].head(5).tolist())}...")

print("\n[ВЫВОД ПО EDA] Данные имеют пропуски (отсутствие наблюдений) и высокую волатильность спроса. ")
print("Модель должна игнорировать пустые остатки при расчете ADS и учитывать сезонность.")
print("="*60 + "\n")
# ==============================================================================
# РАСЧЕТ И ВЫВОД СЕЗОННЫХ КОЭФФИЦИЕНТОВ
# ==============================================================================
print_header("РАСЧЕТ СЕЗОННОСТИ")
print("Что делаем: Считаем, как меняется спрос по месяцам. Если коэффициент > 1 — высокий сезон, < 1 — низкий.")

# Создаем копию истории и извлекаем номер месяца
hist_copy = hist_df.copy()
hist_copy['Month'] = hist_copy['Date'].dt.month

# Берем только дни, когда товар был в наличии (чтобы не искажать спрос дефицитом)
available_hist = hist_copy[hist_copy[STOCK_COL] > 0]

# 1. Считаем МАКРО-ИНДЕКС (по всем SKU вместе)
# Для этого суммируем продажи всех SKU в каждом месяце и делим на среднее
total_sales_by_month = available_hist.groupby('Month')['SalesQty'].sum()
avg_total_sales = total_sales_by_month.mean()
macro_index = (total_sales_by_month / avg_total_sales).rename('Macro_Seasonality_Index')

print("\n[INFO] Общий (макро) индекс сезонности по всем SKU:")
for month, idx in macro_index.items():
    month_name = pd.to_datetime(f"2024-{int(month)}-01").strftime('%B')
    print(f"  {month_name}: {idx:.2f}")

# 2. Считаем ИНДИВИДУАЛЬНЫЙ ИНДЕКС (с защитой от шума)
# Считаем не просто среднее, а сумму продаж за месяц
sku_monthly_sum = available_hist.groupby(['SKU', 'Month'])['SalesQty'].sum().reset_index()
sku_monthly_days = available_hist.groupby(['SKU', 'Month'])['Date'].count().reset_index().rename(columns={'Date': 'DaysCount'})
sku_monthly = sku_monthly_sum.merge(sku_monthly_days, on=['SKU', 'Month'])

# Считаем общее количество продаж по SKU за все время
sku_total = available_hist.groupby('SKU')['SalesQty'].sum().reset_index().rename(columns={'SalesQty': 'TotalSales'})
sku_monthly = sku_monthly.merge(sku_total, on='SKU')

# Считаем индивидуальный коэффициент
sku_monthly['Ind_Seasonality_Index'] = (sku_monthly['SalesQty'] / sku_monthly['TotalSales']) * 12

# Защита от шума: если за месяц было меньше 10 продаж, или за ВСЮ историю было меньше 30 продаж ставим NaN  (нейтрально)
sku_monthly.loc[sku_monthly['SalesQty'] < 10, 'Ind_Seasonality_Index'] = np.nan
low_volume_skus = sku_total[sku_total['TotalSales'] < 30]['SKU'].tolist()
sku_monthly.loc[sku_monthly['SKU'].isin(low_volume_skus), 'Ind_Seasonality_Index'] = np.nan

# Заполняем пропуски (NaN) макро-индексом
# Для этого создаем словарь соответствия месяц -> макро-коэф
macro_dict = macro_index.to_dict()
sku_monthly['Month_Index'] = sku_monthly['Month'].map(macro_dict)

# Если индивидуальный индекс пустой (было мало продаж), берем макро
sku_monthly['Final_Index'] = sku_monthly['Ind_Seasonality_Index'].fillna(sku_monthly['Month_Index'])
sample_skus = sku_df['SKU'].head(5).tolist()
print("\n[INFO] Пример ОЧИЩЕННЫХ индивидуальных коэффициентов (первые 5 SKU):")
sample_df = sku_monthly[sku_monthly['SKU'].isin(sample_skus)].pivot(index='SKU', columns='Month', values='Final_Index').round(2)
print(sample_df)

# ==============================================================================
# БЛОК 2: ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ РАСЧЕТА
# Что делаем: Описываем логику расчета ADS, страхового запаса и самого заказа.
# ==============================================================================
print_header("БЛОК 2: ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ РАСЧЕТА")
print("Что делаем: Описываем логику. Здесь мы считаем ADS (игнорируя дни дефицита),")
print("определяем защитный период (PP) и вычисляем рекомендуемый заказ.")

def get_order_cycle_days(target_date, calendar):
    """Определяет длину цикла (2 или 5 дней) для конкретной даты."""
    row = calendar[calendar['OrderDate'] == target_date]
    if not row.empty:
        return int(row['OrderCycleDays'].iloc[0])
    return 5 # Если даты нет в календаре, берем цикл по умолчанию (5 дней)

def calculate_ads_and_std(hist, as_of_date, window_days=90):
    """
    Считает ADS и стандартное отклонение.
    Игнорирует дни с нулевым остатком (Stockout), чтобы не занижать спрос.
    """
    window_start = as_of_date - timedelta(days=window_days)
    
    # Фильтруем: берем только окно истории и только дни, когда товар был в наличии
    available = hist[(hist['Date'] > window_start) & (hist['Date'] <= as_of_date) & (hist[STOCK_COL] > 0)]
    
    stats = available.groupby('SKU').agg(
        ADS=('SalesQty', 'mean'),
        STD=('SalesQty', 'std')
    ).reset_index()
    
    stats['STD'] = stats['STD'].fillna(0) # Если истории мало, std = 0
    return stats

def calculate_order_logic(target_date, current_stock, supply_data, hist, z_score=1.645, use_seasonality=True):
    """
    Основная логика расчета рекомендуемого заказа.
    Учитывает текущий остаток, открытые поставки и календарь.
    """
    # 1. Защитный период (PP) = Время доставки + Длина цикла
    cycle_days = get_order_cycle_days(target_date, calendar_df)
    pp_days = LEAD_TIME + cycle_days
    target_receipt_date = target_date + timedelta(days=pp_days)
    
    # 2. Получаем ADS и STD на дату заказа
    demand_stats = calculate_ads_and_std(hist, target_date)
    
    # 3. Считаем IP (Inventory Position)
    ip_df = current_stock.copy()
    
    # Добавляем InTransit и OpenPO, которые придут ДО конца защитного периода
    if not supply_data.empty:
        relevant_supply = supply_df[
            (supply_df['ExpectedReceiptDate'] <= target_receipt_date) & 
            (supply_df['ExpectedReceiptDate'] > target_date)
        ]
        supply_sum = relevant_supply.groupby('SKU')['Qty'].sum().reset_index().rename(columns={'Qty': 'SupplyQty'})
        ip_df = ip_df.merge(supply_sum, on='SKU', how='left').fillna({'SupplyQty': 0})
    else:
        ip_df['SupplyQty'] = 0
        
    ip_df['IP'] = ip_df['CurrentStock'] + ip_df['SupplyQty']
    
    # 4. Объединяем с ADS и считаем SS (Страховой запас) 
    df = ip_df.merge(demand_stats, on='SKU', how='left').fillna({'ADS': 0, 'STD': 0})
    
    df['SS'] = np.where(df['STD'] == 0, 0, z_score * df['STD'] * np.sqrt(pp_days))

    # 5. Считаем Целевой запас (S) с учетом сезонности
    if use_seasonality:
        target_month = target_date.month
        season_df = sku_monthly[sku_monthly['Month'] == target_month][['SKU', 'Final_Index']].rename(columns={'Final_Index': 'SeasonCoef'})
        df = df.merge(season_df, on='SKU', how='left')
        df['SeasonCoef'] = df['SeasonCoef'].fillna(1.0)
        df['S'] = (df['ADS'] * df['SeasonCoef']) * pp_days + df['SS']
    else:
        df['S'] = df['ADS'] * pp_days + df['SS']

    # 6. Считаем заказ
    df['RecommendedOrder'] = np.ceil(df['S'] - df['IP']).clip(lower=0).astype(int)
    df.loc[df['ADS'] == 0, 'RecommendedOrder'] = 0 # Если нет спроса, не заказываем
    
    return df[['SKU', 'IP', 'ADS', 'SS', 'S', 'RecommendedOrder']]

# --- ДЕМОНСТРАЦИЯ РАБОТЫ ФУНКЦИИ ---
print("\n[ДЕМО] Выводим первые 5 строк таблицы, которую возвращает функция (на дату 19.09.2025):")
# Для демо берем реальные текущие остатки и реальные открытые поставки
demo_stock = sku_df[['SKU', 'CurrentStock']].copy()
demo_result = calculate_order_logic(
    target_date=AS_OF_DATE,
    current_stock=demo_stock,
    supply_data=supply_df,
    hist=hist_df,
    z_score=1.645,
    use_seasonality=True
)
print(demo_result.head())

# ==============================================================================
# БЛОК 3: СИМУЛЯЦИЯ (БЭКТЕСТИНГ)
# Что делаем: "Проживаем" время заново, принимая решения только на основе прошлого.
# ==============================================================================
print_header("БЛОК 3: СИМУЛЯЦИЯ (БЭКТЕСТИНГ)")
print("Что делаем: 'Проживаем' время заново. На каждую дату смотрим остатки,")
print("вычитаем реальные продажи, ждем прихода заказов и принимаем новые решения.")
print("В конце функция возвращает таблицу с метриками эффективности (FillRate, AvgInventory и т.д.).")

def run_backtest_simulation(start_date, end_date, initial_stock_dict, z_score, is_baseline=False, baseline_window=7, use_seasonality=True):
    """
    Симуляция работы склада.
    is_baseline=True включает простую базовую модель (Заказ = Продажи за прошлую неделю)
    """
    dates = pd.date_range(start_date, end_date)
    inventory = initial_stock_dict.copy()
    pipeline = [] # Очередь заказов, которые еще едут
    
    metrics = {
        'lost_sales': {sku: 0 for sku in sku_df['SKU']},
        'stockout_days': {sku: 0 for sku in sku_df['SKU']},
        'inventory_levels': {sku: [] for sku in sku_df['SKU']},
    }
    
    # Берем только те даты заказов, которые попадают в наш период симуляции
    order_dates = calendar_df[(calendar_df['OrderDate'] >= start_date) & (calendar_df['OrderDate'] <= end_date)]['OrderDate'].tolist()

        # 1. Приход поставок (если дата прибытия наступила)
    for current_date in dates:
    
        arrived = [p for p in pipeline if p['arrival_date'] <= current_date]
        for p in arrived:
            inventory[p['sku']] += p['qty'] # Добавляем к остатку
        pipeline = [p for p in pipeline if p['arrival_date'] > current_date]
        
        # 2. Продажи (берем фактические из истории) 
        daily_sales = hist_df[hist_df['Date'] == current_date].set_index('SKU')['SalesQty']
        
        for sku in sku_df['SKU']:
            sales = daily_sales.get(sku, 0) # Факт из истории
            available = inventory.get(sku, 0)
            
            if available >= sales:
                inventory[sku] = available - sales # ... вычитаем продажи из остатка ...
            else:
                # Дефицит (Lost Sales) - спрос превысил остаток
                lost = sales - available
                metrics['lost_sales'][sku] += lost
                inventory[sku] = 0
                metrics['stockout_days'][sku] += 1
                
            
            metrics['inventory_levels'][sku].append(inventory.get(sku, 0))
            
        # 3. Принятие решения о заказе (если сегодня день заказа по календарю)
        if current_date in order_dates: 
            # Берем текущий ВИРТУАЛЬНЫЙ остаток
            curr_stock_df = pd.DataFrame({'SKU': list(inventory.keys()), 'CurrentStock': list(inventory.values())})
            
            # Берем историю продаж ТОЛЬКО ДО СЕГОДНЯШНЕГО ДНЯ (No Look-Ahead Bias)
            hist_past = hist_df[hist_df['Date'] < current_date]
            
            if is_baseline:
                # БАЗОВАЯ МОДЕЛЬ: Заказываем столько, сколько продали за последние 7 дней
                last_week_sales = hist_past[hist_past['Date'] > (current_date - timedelta(days=baseline_window))].groupby('SKU')['SalesQty'].sum().reset_index()
                order_recs = curr_stock_df.merge(last_week_sales, on='SKU', how='left').fillna({'SalesQty': 0})
                order_recs['RecommendedOrder'] = order_recs['SalesQty'].astype(int) # Просто копируем спрос
            else:
                # НАША МОДЕЛЬ
                order_recs = calculate_order_logic(
                    target_date=current_date,
                    current_stock=curr_stock_df,
                    supply_data=pd.DataFrame(),
                    hist=hist_past,
                    z_score=z_score,
                    use_seasonality=use_seasonality 
                )

                
            # Кладем наш виртуальный заказ в "трубу" (приедет через 10 дней)
            for _, row in order_recs.iterrows():
                if row['RecommendedOrder'] > 0:
                    pipeline.append({
                        'sku': row['SKU'],
                        'qty': row['RecommendedOrder'],
                        'arrival_date': current_date + timedelta(days=LEAD_TIME)
                    })

    # Сбор итоговых метрик по каждому SKU
    results = []
    for sku in sku_df['SKU']:
        total_sold = hist_df[(hist_df['SKU'] == sku) & (hist_df['Date'] >= start_date) & (hist_df['Date'] <= end_date)]['SalesQty'].sum()
        lost = metrics['lost_sales'][sku]
        total_demand = total_sold + lost
        
        fill_rate = total_sold / total_demand if total_demand > 0 else 1.0
        inv_levels = metrics['inventory_levels'][sku]
        
        results.append({
            'SKU': sku,
            'FillRate': fill_rate,
            'AvgInventory': np.mean(inv_levels) if inv_levels else 0,
            'MaxInventory': np.max(inv_levels) if inv_levels else 0,
            'StockoutDays': metrics['stockout_days'][sku],
            'TotalLostSales': lost
        })
        
    return pd.DataFrame(results)



# ==============================================================================
# БЛОК 4: НАСТРОЙКА ПАРАМЕТРОВ (TRAIN)
# Что делаем: Ищем оптимальный Z-score на исторических данных.
# ==============================================================================
print_header("БЛОК 4: НАСТРОЙКА ПАРАМЕТРОВ (TRAIN: 01.01.2025 - 31.05.2025)")
print("Что делаем: Ищем лучший коэффициент Z (уровень сервиса), который дает")
print("хороший Fill Rate при минимальном среднем запасе. Модель обучается на этих данных.")


start_train = pd.to_datetime("2025-01-01")
end_train = pd.to_datetime("2025-05-31")
prev_day_train = start_train - timedelta(days=1)
initial_stock_train = dict(zip(hist_df[hist_df['Date'] == prev_day_train]['SKU'], hist_df[hist_df['Date'] == prev_day_train][STOCK_COL]))

# Тестируем разные Z-коэффициенты (уровни сервиса)
print("\n[INFO] Перебираем Z-коэффициенты:")
for z in [1.28, 1.645, 2.33]:
    res = run_backtest_simulation(start_train, end_train, initial_stock_train, z_score=z)
    print(f"Z-score: {z} | Средний Fill Rate: {res['FillRate'].mean():.2%} | Средний запас: {res['AvgInventory'].mean():.1f}")

# Допустим, мы выбрали Z = 1.645 (95% уровень сервиса)
BEST_Z = 1.645
print(f"\n[РЕШЕНИЕ] Выбираем Z = {BEST_Z} (баланс между сервисом и запасом).")


# ==============================================================================
# БЛОК 5: КОНТРОЛЬНЫЕ ДАТЫ (TEST)
# Что делаем: Проверяем модель на данных, которые она "не видела" при настройке.
# ==============================================================================
print_header("БЛОК 5: ПРОВЕРКА НА КОНТРОЛЬНЫХ ДАТАХ (TEST: 01.06.2025 - 18.09.2025)")
print("Что делаем: Проверяем модель на данных, которые она 'не видела' при обучении.")
print("Сравниваем нашу модель (+с сезонностью) с базовыми (заказ на основе продаж за 7,15,30).")


start_test = pd.to_datetime("2025-06-01")
end_test = pd.to_datetime("2025-09-18")
prev_day_test = start_test - timedelta(days=1)
initial_stock_test = dict(zip(hist_df[hist_df['Date'] == prev_day_test]['SKU'], hist_df[hist_df['Date'] == prev_day_test][STOCK_COL]))

# 1. Наша модель БЕЗ сезонности
test_results_no_season = run_backtest_simulation(
    start_test, end_test, initial_stock_test, 
    z_score=BEST_Z, is_baseline=False, use_seasonality=False
)

# 2. Наша модель С сезонностью
test_results_with_season = run_backtest_simulation(
    start_test, end_test, initial_stock_test, 
    z_score=BEST_Z, is_baseline=False, use_seasonality=True
)

# 3. Базовые модели 
test_results_base_7 = run_backtest_simulation(
    start_test, end_test, initial_stock_test, 
    z_score=BEST_Z, is_baseline=True, baseline_window=7
)
test_results_base_15 = run_backtest_simulation(
    start_test, end_test, initial_stock_test, 
    z_score=BEST_Z, is_baseline=True, baseline_window=15
)
test_results_base_30 = run_backtest_simulation(
    start_test, end_test, initial_stock_test, 
    z_score=BEST_Z, is_baseline=True, baseline_window=30
)
print(f"{'Модель':<35} | {'Fill Rate':<10} | {'Ср. запас':<10} | {'Макс. запас':<12} | {'Дефицит (дн)':<12} | {'Упущено (шт)':<10}")
print("-" * 85)

# Вывод нашей модели без сезонности
m1 = test_results_no_season
print(f"{'НАША МОДЕЛЬ (без сезонности)':<35} | {m1['FillRate'].mean():>9.2%} | {m1['AvgInventory'].mean():>10.1f} | {m1['MaxInventory'].max():>12.0f} | {m1['StockoutDays'].sum():>12.0f} | {m1['TotalLostSales'].sum():>10.0f}")

# Вывод нашей модели с сезонностью
m2 = test_results_with_season
print(f"{'НАША МОДЕЛЬ (с сезонностью)':<35} | {m2['FillRate'].mean():>9.2%} | {m2['AvgInventory'].mean():>10.1f} | {m2['MaxInventory'].max():>12.0f} | {m2['StockoutDays'].sum():>12.0f} | {m2['TotalLostSales'].sum():>10.0f}")

# Вывод базовой модели (7 дней)
m3 = test_results_base_7
print(f"{'БАЗОВАЯ (окно 7 дней)':<35} | {m3['FillRate'].mean():>9.2%} | {m3['AvgInventory'].mean():>10.1f} | {m3['MaxInventory'].max():>12.0f} | {m3['StockoutDays'].sum():>12.0f} | {m3['TotalLostSales'].sum():>10.0f}")

# Вывод базовой модели (15 дней)
m4 = test_results_base_15
print(f"{'БАЗОВАЯ (окно 15 дней)':<35} | {m4['FillRate'].mean():>9.2%} | {m4['AvgInventory'].mean():>10.1f} | {m4['MaxInventory'].max():>12.0f} | {m4['StockoutDays'].sum():>12.0f} | {m4['TotalLostSales'].sum():>10.0f}")

# Вывод базовой модели (30 дней)
m5 = test_results_base_30
print(f"{'БАЗОВАЯ (окно 30 дней)':<35} | {m5['FillRate'].mean():>9.2%} | {m5['AvgInventory'].mean():>10.1f} | {m5['MaxInventory'].max():>12.0f} | {m5['StockoutDays'].sum():>12.0f} | {m5['TotalLostSales'].sum():>10.0f}")
print("-" * 85)

# Вывод разницы между нашей моделью с и без сезонности
diff_fill = (m2['FillRate'].mean() - m1['FillRate'].mean()) * 100
diff_inv = m2['AvgInventory'].mean() - m1['AvgInventory'].mean()
diff_stockout = m2['StockoutDays'].sum() - m1['StockoutDays'].sum()
diff_lost = m2['TotalLostSales'].sum() - m1['TotalLostSales'].sum()

print(f"\n[ЭФФЕКТ ОТ СЕЗОННОСТИ]")
print(f"  Прирост Fill Rate:      {diff_fill:+.2f}%")
print(f"  Изменение ср. запаса:   {diff_inv:+.1f} шт.")
print(f"  Сокращение дефицита:    {diff_stockout:+.0f} дн.")
print(f"  Сокращение упущенных:   {diff_lost:+.0f} шт.")

# ==============================================================================
# ДОПОЛНИТЕЛЬНО: СНИМОК (SNAPSHOT) НА КОНКРЕТНЫЕ КОНТРОЛЬНЫЕ ДАТЫ
# ==============================================================================
print("\n" + "="*60)
print("ПРИМЕР РАСЧЕТА НА КОНКРЕТНЫЕ ПОСЛЕДОВАТЕЛЬНЫЕ ДАТЫ (Июнь 2025)")
print("="*60)

# Выберем 5 последовательных пятниц июня
control_dates = [
    pd.to_datetime("2025-06-06"),
    pd.to_datetime("2025-06-13"),
    pd.to_datetime("2025-06-20"),
    pd.to_datetime("2025-06-27"),
    pd.to_datetime("2025-07-04")
]

# Возьмем один SKU для наглядности (например, TEST-0025, который мы разбирали)
target_sku = "TEST-0025"

print(f"Анализ для SKU: {target_sku}")
print(f"{'Дата':<12} | {'Остаток (IP)':<12} | {'ADS':<6} | {'Целевой S':<10} | {'Заказ':<6}")
print("-" * 60)

for dt in control_dates:
    # Берем историю до этой даты
    hist_past = hist_df[hist_df['Date'] < dt]
    
    # Берем остаток на конец предыдущего дня (как мы делали в симуляции)
    prev_day = dt - timedelta(days=1)
    stock_row = hist_df[(hist_df['Date'] == prev_day) & (hist_df['SKU'] == target_sku)]
    
    if not stock_row.empty:
        current_st = stock_row[STOCK_COL].values[0]
        
        # Создаем df для расчета
        curr_stock_df = pd.DataFrame({'SKU': [target_sku], 'CurrentStock': [current_st]})
        
        # Считаем заказ (БЕЗ сезонности для простоты, или можно с ней)
        order_rec = calculate_order_logic(
            target_date=dt, 
            current_stock=curr_stock_df, 
            supply_data=pd.DataFrame(), 
            hist=hist_past, 
            z_score=BEST_Z,
            use_seasonality=True # Включаем сезонность
        )
        
        row = order_rec.iloc[0]
        print(f"{dt.date()}   | {row['IP']:<12.0f} | {row['ADS']:<6.2f} | {row['S']:<10.1f} | {row['RecommendedOrder']:<6.0f}")
    else:
        print(f"{dt.date()}   | Нет данных в истории")
# ==============================================================================
# БЛОК 6: ФИНАЛЬНЫЙ РАСЧЕТ НА 19.09.2025
# Что делаем: Считаем заказ на дату расчета, используя ВСЕ доступные данные (включая OpenSupply).
# ==============================================================================

print_header("БЛОК 6: ФИНАЛЬНЫЙ РАСЧЕТ НА 19.09.2025")
print("Что делаем: Считаем заказ на дату расчета, используя все доступные данные.")
print("В отличие от бэктеста, здесь мы уже ЗНАЕМ про открытые поставки (OpenSupply).")


current_stock_final = sku_df[['SKU', 'CurrentStock']].copy()

final_order = calculate_order_logic(
    target_date=AS_OF_DATE,
    current_stock=current_stock_final,
    supply_data=supply_df, # Здесь уже передаем реальные открытые поставки!
    hist=hist_df,
    z_score=BEST_Z,
    use_seasonality=True
)

print(f"\n[РЕЗУЛЬТАТ] Общая сумма заказа: {final_order['RecommendedOrder'].sum()} шт.")
print("\n[INFO] Шапка итоговой таблицы (первые 5 строк):")
print(final_order.head())


# Сохраняем результат в Excel
final_order[['SKU', 'RecommendedOrder']].to_excel("final_order_19_09_2025.xlsx", index=False)
print("\nФайл 'final_order_19_09_2025.xlsx' успешно сохранен.")