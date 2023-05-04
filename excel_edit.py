import pandas as pd
import openpyxl
import os

ranges = [[0, 10, 0], [10, 50, 0.5], [50, 80, 1], [80, 130, 1.5], [130, 170, 2], [170, 210, 2.5], [210, 260, 3], [260, 320, 3.5], [320, 365, 4], [365, 410, 4.5], [410, 460, 5], [460, 500, 5.5], [500, 540, 6], [540, 590, 6.5], [590, 630, 7], [630, 680, 7.5], [680, 720, 8], [720, 760, 8.5], [760, 800, 9], [800, 860, 9.5], [860, 910, 10], [910, 955, 10.5], [955, 1005, 11], [1005, 1065, 11.5], [1065, 1095, 12], [1095, 1145, 12.5], [1145, 1185, 13], [1185, 1230, 13.5], [1230, 1290, 14], [
    1290, 1350, 14.5], [1350, 1390, 15], [1390, 1430, 15.5], [1430, 1490, 16], [1490, 1535, 16.5], [1535, 1590, 17], [1590, 1640, 17.5], [1640, 1690, 18], [1690, 1740, 18.5], [1740, 1785, 19], [1785, 1810, 19.5], [1810, 1860, 20], [1860, 1900, 20.5], [1900, 1940, 21], [1940, 1985, 21.5], [1985, 2015, 22], [2015, 2085, 22.5], [2085, 2115, 23], [2115, 2195, 23.5], [2195, 2240, 24], [2240, 2295, 24.5], [2295, 2330, 25], [2330, 2390, 25.5], [2390, 2430, 26], [2430, 2475, 26.5], [2475, 2510, 27], [2510, 2750, 27.5], [2750, 2900, 28], [2900, 3100, 29], [3100, 3300, 30]]

# Changing row if
def if_poland(row):
    if row["Państwo docelowe"] == "Polska":
        return "Nie"

# Count how many hours based on ranges
def estimate_hours(row):
    try:
        distance = int(row["Dystans [km]"])
    except ValueError:
        return ("Brak danych o liczbie kilometrów")
    for elem in ranges:
        if distance > elem[0] and distance <= elem[1]:
            return (elem[2])

    return distance

# Creates excel file wtih dataframe and driver's name
def create_excel_table(new_df, driver, output_path):
    string = driver
    string = ''.join(e for e in string if e.isalnum())
    path = output_path + "\\" + string + ".xlsx"
    writer = pd.ExcelWriter(path)
    new_df.drop('Samochód', inplace=True, axis=1)
    new_df.to_excel(writer, sheet_name='Raport', startrow=1,
                    startcol=0, index=False, na_rep='')
    worksheet = writer.sheets['Raport']
    worksheet.write(0, 0, string)
    for column in new_df:
        column_length = max(new_df[column].astype(
            str).map(len).max(), len(column))
        col_idx = new_df.columns.get_loc(column)
        worksheet.set_column(col_idx, col_idx, column_length)
    writer.save()


# Cleaning data and separate one drivers from list
def data_clean(input_path, output_path):
    df = pd.read_excel(input_path, engine='openpyxl')
    df.drop('Zużyty gaz [kg]', inplace=True, axis=1)
    df.drop('Zużyte paliwo [l]', inplace=True, axis=1)
    df.drop('Dystans (GPS) [km]', inplace=True, axis=1)
    df.drop('Dystans (CAN) [km]', inplace=True, axis=1)
    df.drop('Czas przebywania', inplace=True, axis=1)
    df.drop('Nr rej.', inplace=True, axis=1)
    df.drop('Czas jazdy', inplace=True, axis=1)
    df.dropna(
        axis=0,
        how='all',
        subset=["Czas wjazdu", "Czas wyjazdu"],
        inplace=True
    )
    df['Czas wjazdu'] = df['Czas wjazdu'].dt.strftime('%Y-%m-%d %H:%M')
    df['Czas wyjazdu'] = df['Czas wyjazdu'].dt.strftime('%Y-%m-%d %H:%M')
    df[['Data wjazdu', 'Godzina wjazdu']] = df['Czas wjazdu'].astype(
        str).str.split(' ', 1, expand=True)
    df[['Data wyjazdu', 'Godzina wyjazdu']] = df['Czas wyjazdu'].astype(
        str).str.split(' ', 1, expand=True)
    df.drop('Czas wjazdu', inplace=True, axis=1)
    df.drop('Czas wyjazdu', inplace=True, axis=1)
    df["Państwo docelowe"] = df["Kraj"]

    df['Uwzględnij jako delegowanie'] = df.apply(
        lambda row: if_poland(row), axis=1)
    df["Ilość godzin pracy"] = df.apply(
        lambda row: estimate_hours(row), axis=1)
    df.drop('Dystans [km]', inplace=True, axis=1)
    df = df.fillna("")
    drivers = (pd.unique(df["Samochód"]))
    for driver in drivers:
        new_df = df[df["Samochód"] == driver].copy()
        create_excel_table(new_df, driver, output_path)
