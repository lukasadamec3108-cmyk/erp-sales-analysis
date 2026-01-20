import pandas as pd

# načtení ERP dat z Excelu
file_path = "ERP_SIMULATION.xlsx"
df = pd.read_excel(file_path)

# základní kontrola
print("Data byla úspěšně načtena")
print(df.head())

print("\nINFO O DATECH")
print(df.info())

# Základní přehled dat
print("\nZÁKLADNÍ STATISTIKY")
print(df.describe())

# Celkové tržby (KPI)
total_revenue = df["Revenue"].sum()
print("\nCelkové tržby:", total_revenue)

# Počet dokladů
order_count = df["Document"].nunique()
print("Počet dokladů:", order_count)

# Průměrná tržba na doklad
avg_revenue = total_revenue / order_count
print("Průměrná tržba na doklad:", avg_revenue)

# Tržby podle CostCenter (GROUP BY)
revenue_by_cc = df.groupby("CostCenter")["Revenue"].sum()
print("\nTržby podle CostCenter:")
print(revenue_by_cc)

# Tržby pouze pro CZ
cz_revenue = df[df["CostCenter"] == "CZ"]["Revenue"].sum()
print("Tržby pro CZ:", cz_revenue)

# Převod Date na skutečné datum
df["Date"] = pd.to_datetime(df["Date"])

# Seřazení dat
df = df.sort_values("Date")

# Celkové YTD tržby
ytd_revenue = df["Revenue"].sum()
print("\nYTD tržby:", ytd_revenue)

# poslední měsíc v datech:
last_month = df["Date"].dt.to_period("M").max()
print("Poslední měsíc v datech:", last_month)

# MTD tržby
mtd_revenue = df[
    df["Date"].dt.to_period("M") == last_month
]["Revenue"].sum()

print("MTD tržby:", mtd_revenue)

# MTD tržby pouze pro CZ
cz_mtd_revenue = df[
    (df["Date"].dt.to_period("M") == last_month) &
    (df["CostCenter"] == "CZ")
]["Revenue"].sum()

print("MTD tržby pro CZ:", cz_mtd_revenue)

# Výpočet 95. percentilu
percentile_95 = df["Revenue"].quantile(0.95)
print("95. percentil tržeb:", percentile_95)

# Vytvoření flag sloupce
df["Above_95_Percentile"] = df["Revenue"] > percentile_95
print(df)

# TABULKA POUZE EXTRÉMNÍCH DOKLADŮ
outliers = df[df["Above_95_Percentile"]]
print("\nExtrémní doklady:")
print(outliers)

# % tržeb tvořící extrémní doklady
outlier_revenue = outliers["Revenue"].sum()
share_outliers = outlier_revenue / total_revenue * 100

print("\nPodíl extrémních tržeb (%):", round(share_outliers, 2))

# vyfiltrujeme pouze extrémní doklady
outliers = df[df["Above_95_Percentile"] == True]

print("Extrémní doklady:")
print(outliers)

# export do Excelu
outliers.to_excel("extremni_doklady.xlsx", index=False)

print("Report byl uložen jako extremni_doklady.xlsx")

# agregace KPI podle CostCenter
kpi_cc = df.groupby("CostCenter").agg(
    Pocet_dokladu=("Document", "count"),
    Celkove_trzby=("Revenue", "sum"),
    Prumerna_trzba=("Revenue", "mean"),
    Max_trzba=("Revenue", "max")
)

print("\nKPI podle CostCenter:")

# dataset bez extrémních dokladů
df_clean = df[df["Above_95_Percentile"] == False]

print("\nData bez extrémních dokladů:")
print(df_clean)

# KPI PODLE COSTCENTER – BEZ EXTRÉMŮ
kpi_clean_cc = df_clean.groupby("CostCenter").agg(
    Pocet_dokladu=("Document", "count"),
    Celkove_trzby=("Revenue", "sum"),
    Prumerna_trzba=("Revenue", "mean"),
    Max_trzba=("Revenue", "max")
)

print("\nKPI bez extrémů podle CostCenter:")
print(kpi_clean_cc)

# DEFINICE AKTUÁLNÍHO DATA
today = pd.to_datetime("2024-02-15")

# FUNKCE PRO MTD / YTD
def calculate_mtd_ytd(data, date_col, value_col, today):
    data = data.copy()
    data[date_col] = pd.to_datetime(data[date_col])

    mtd = data[
        (data[date_col].dt.year == today.year) &
        (data[date_col].dt.month == today.month) &
        (data[date_col] <= today)
    ][value_col].sum()

    ytd = data[
        (data[date_col].dt.year == today.year) &
        (data[date_col] <= today)
    ][value_col].sum()

    return mtd, ytd

# VÝPOČET – RAW DATA
mtd_raw, ytd_raw = calculate_mtd_ytd(df, "Date", "Revenue", today)

# VÝPOČET – CLEAN DATA
mtd_clean, ytd_clean = calculate_mtd_ytd(df_clean, "Date", "Revenue", today)

# VÝSTUP
print("\nMTD / YTD srovnání:")
print(f"RAW MTD: {mtd_raw}, RAW YTD: {ytd_raw}")
print(f"CLEAN MTD: {mtd_clean}, CLEAN YTD: {ytd_clean}")

print("\n--- SANITY CHECKS ---")

# 1️⃣ prázdné hodnoty
print("Chybějící hodnoty:")
print(df.isnull().sum())

# 2️⃣ datové typy
print("\nDatové typy:")
print(df.dtypes)

# 3️⃣ záporné tržby
negative_revenue = df[df["Revenue"] < 0]
print("\nZáporné tržby:")
print(negative_revenue)

# tržby == 0
zero_revenue = df[df["Revenue"] == 0]

# neznámé cost center
valid_cc = ["CZ", "SK"]
invalid_cc = df[~df["CostCenter"].isin(valid_cc)]

# datum v budoucnosti
future_dates = df[df["Date"] > today]

print("\nNulové tržby:")
print(zero_revenue)

print("\nNeplatné CostCenter:")
print(invalid_cc)

print("\nBudoucí datum:")
print(future_dates)

df_clean = df.copy()

# odstraníme nulové a záporné tržby
df_clean = df_clean[df_clean["Revenue"] > 0]

# odstraníme neznámá cost center
df_clean = df_clean[df_clean["CostCenter"].isin(valid_cc)]

# odstraníme budoucí data
df_clean = df_clean[df_clean["Date"] <= today]

print("\nPočet záznamů RAW:", len(df))
print("Počet záznamů CLEAN:", len(df_clean))