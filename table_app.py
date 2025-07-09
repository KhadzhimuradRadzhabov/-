import pandas as pd
import numpy as np
import streamlit as st
import io
from collections import OrderedDict

input_file = r"Основная таблица.csv"
table = pd.read_csv(input_file)
table = table.drop(columns=['Unnamed: 0'])

st.title("Фильтр по МНН и классификациям")
buffer = io.BytesIO()

# фильтр
mnn_selection = st.multiselect("Выберите МНН", sorted(table['Molecule'].unique()))
class1_selection = st.multiselect("Выберите New Form Classification Lev 1", sorted(table['New Form Classification Lev 1'].unique()))
class2_selection = st.multiselect("Выберите New Form Classification Lev 2", sorted(table['New Form Classification Lev 2'].unique()))
class3_selection = st.multiselect("Выберите New Form Classification Lev 3", sorted(table['New Form Classification Lev 3'].unique()))
ephmra1_selection = st.multiselect("Выберите EphMRA1", sorted(table['EphMRA1'].unique()))
ephmra2_selection = st.multiselect("Выберите EphMRA2", sorted(table['EphMRA2'].unique()))
ephmra3_selection = st.multiselect("Выберите EphMRA3", sorted(table['EphMRA3'].unique()))
atc1_selection = st.multiselect("Выберите ATC1", sorted(table['ATC1'].dropna().astype(str).unique()))
atc2_selection = st.multiselect("Выберите ATC2", sorted(table['ATC2'].dropna().astype(str).unique()))
atc3_selection = st.multiselect("Выберите ATC3", sorted(table['ATC3'].dropna().astype(str).unique()))

# собираем таблицу и выводим
if mnn_selection:
    st.subheader("Результат по выбранным МНН")
    
    df_selected = table[table['Molecule'].isin(mnn_selection)]
    aggregated_rows = []

    for mnn in df_selected['Molecule'].unique():
        df_mnn = df_selected[df_selected['Molecule'] == mnn]

        if len(df_mnn) > 1:
            row_max_players = df_mnn.loc[df_mnn["Кол-во игроков (МНН+NFC)"].idxmax()]
        
            # Вычисления
            sum_19_rub = df_mnn["Сумма 19, М Руб"].sum()
            sum_24_rub = df_mnn["Сумма 24, М Руб"].sum()
            sum_19_up = df_mnn["Сумма 19, тыс уп"].sum()
            sum_24_up = df_mnn["Сумма 24, тыс уп"].sum()

            # Расчёт CAGR
            cagr_rub = ((sum_24_rub / sum_19_rub) ** (1 / 4) - 1) if sum_19_rub > 0 else None
            cagr_up = ((sum_24_up / sum_19_up) ** (1 / 4) - 1) if sum_19_up > 0 else None

            aggregated = OrderedDict([
                ("Molecule", mnn),
                ("New Form Classification Lev 1", row_max_players["New Form Classification Lev 1"]),
                ("New Form Classification Lev 2", row_max_players["New Form Classification Lev 2"]),
                ("New Form Classification Lev 3", row_max_players["New Form Classification Lev 3"]),
                ("EphMRA1", row_max_players["EphMRA1"]),
                ("EphMRA2", row_max_players["EphMRA2"]),
                ("EphMRA3", row_max_players["EphMRA3"]),
                ("ATC1", row_max_players["ATC1"]),
                ("ATC2", row_max_players["ATC2"]),
                ("ATC3", row_max_players["ATC3"]),
                ("RX/OTC", row_max_players["RX/OTC"]),
                ("Essential Drug List", row_max_players["Essential Drug List"]),
                ("ТОП-1 Бренд (руб.)", row_max_players["ТОП-1 Бренд (руб.)"]),
                ("Доля ТОП-1 Бренда (руб.)", row_max_players["Доля ТОП-1 Бренда (руб.)"]),
                ("ТОП-1 Бренд (уп.)", row_max_players["ТОП-1 Бренд (уп.)"]),
                ("Доля ТОП-1 Бренда (уп.)", row_max_players["Доля ТОП-1 Бренда (уп.)"]),
                ("Разные лидеры?", row_max_players["Разные лидеры?"]),
                ("Кол-во игроков (МНН+NFC)", df_mnn["Кол-во игроков (МНН+NFC)"].sum()),
                ("Сумма 19, М Руб", sum_19_rub),
                ("Сумма 23, М Руб", df_mnn["Сумма 23, М Руб"].sum()),
                ("Сумма 24, М Руб", sum_24_rub),
                ("Прирост 24, М руб", df_mnn["Прирост 24, М руб"].sum()),
                ("CAGR 5Y, руб", "{:.1%}".format(cagr_rub) if cagr_rub is not None else None),
                ("Сумма 19, тыс уп", sum_19_up),
                ("Сумма 23, тыс уп", df_mnn["Сумма 23, тыс уп"].sum()),
                ("Сумма 24, тыс уп", sum_24_up),
                ("Прирост 24, тыс уп", df_mnn["Прирост 24, тыс уп"].sum()),
                ("CAGR 5Y, уп", "{:.1%}".format(cagr_up) if cagr_up is not None else None),
            ])

            aggregated_rows.append(aggregated)

        else:
            row = df_mnn.iloc[0].to_dict()
            aggregated_rows.append(row)

    final_df = pd.DataFrame(aggregated_rows)
    final_df["CAGR 5Y, руб"] = final_df["CAGR 5Y, руб"].apply(lambda x: "{:.1%}".format(x) if pd.notnull(x) else None)
    final_df["CAGR 5Y, уп"] = final_df["CAGR 5Y, уп"].apply(lambda x: "{:.1%}".format(x) if pd.notnull(x) else None)
    st.dataframe(final_df)

    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
        # Write each dataframe to a different worksheet.
        table[table['Molecule'].isin(mnn_selection)].reset_index(drop=True).to_excel(writer, sheet_name='Sheet1', index=False)
        writer.close()

        download = st.download_button(
            label="Download as Excel",
            data=buffer,
            file_name='Сводная таблица.xlsx',
            #mime='application/vnd.ms-excel'
        )

elif class1_selection or class2_selection or class3_selection or ephmra1_selection or ephmra2_selection or ephmra3_selection or atc1_selection or atc2_selection or atc3_selection:
    df = table.copy()
    if class1_selection:
        df = df[df['New Form Classification Lev 1'].isin(class1_selection)]
    if class2_selection:
        df = df[df['New Form Classification Lev 2'].isin(class2_selection)]
    if class3_selection:
        df = df[df['New Form Classification Lev 3'].isin(class3_selection)]
    if ephmra1_selection:
        df = df[df['EphMRA1'].isin(ephmra1_selection)]
    if ephmra2_selection:
        df = df[df['EphMRA2'].isin(ephmra2_selection)]
    if ephmra3_selection:
        df = df[df['EphMRA3'].isin(ephmra3_selection)]
    if atc1_selection:
        df = df[df['ATC1'].isin(atc1_selection)]
    if atc2_selection:
        df = df[df['ATC2'].isin(atc2_selection)]
    if atc3_selection:
        df = df[df['ATC3'].isin(atc3_selection)]
    
    group_cols = []
    if class1_selection:
        group_cols.append('New Form Classification Lev 1')
    if class2_selection:
        group_cols.append('New Form Classification Lev 2')
    if class3_selection:
        group_cols.append('New Form Classification Lev 3')
    if ephmra1_selection:
        group_cols.append('EphMRA1')
    if ephmra2_selection:
        group_cols.append('EphMRA2')    
    if ephmra3_selection:
        group_cols.append('EphMRA3')
    if atc1_selection:
        group_cols.append('ATC1')
    if atc2_selection:
        group_cols.append('ATC2')
    if atc3_selection:
        group_cols.append('ATC3')

    
    agg_df = df.groupby(group_cols)[[
        'Кол-во игроков (МНН+NFC)',
        'Сумма 23, М Руб',
        'Сумма 24, М Руб'
    ]].sum().reset_index()
    agg_df['Прирост 24, М руб'] = agg_df['Сумма 24, М Руб'] - agg_df['Сумма 23, М Руб']

    
    # Группируем кагр по руб
    base = df.groupby(group_cols)["Сумма 19, М Руб"].sum().reset_index(name="base_19")
    future = df.groupby(group_cols)["Сумма 24, М Руб"].sum().reset_index(name="future_24")
    cagr_df = pd.merge(base, future, on=group_cols, how="outer")

    # считаем CAGR по руб
    agg_df["CAGR 5Y, руб"] = np.where(
        (cagr_df["base_19"] > 0) & (cagr_df["future_24"] >= 0),
        round((cagr_df["future_24"] / cagr_df["base_19"]) ** (1/4) - 1, 2),
        np.nan
    )
    
    agg_sum_23up = df.groupby(group_cols)['Сумма 23, тыс уп'].sum().reset_index()
    agg_df = pd.merge(agg_df, agg_sum_23up, on=group_cols, how='left') 
    
    agg_sum_24up = df.groupby(group_cols)['Сумма 24, тыс уп'].sum().reset_index()
    agg_df = pd.merge(agg_df, agg_sum_24up, on=group_cols, how='left')
    
    agg_df['Прирост 24, тыс уп'] = agg_df['Сумма 24, тыс уп'] - agg_df['Сумма 23, тыс уп']

    # Группируем кагр по уп
    baseu = df.groupby(group_cols)["Сумма 19, тыс уп"].sum().reset_index(name="base_19")
    futureu = df.groupby(group_cols)["Сумма 24, тыс уп"].sum().reset_index(name="future_24")
    cagru_df = pd.merge(baseu, futureu, on=group_cols, how="outer")

    #считаем CAGR по уп
    agg_df["CAGR 5Y, уп"] = np.where(
        (cagru_df["base_19"] > 0) & (cagru_df["future_24"] >= 0), 
        round((cagru_df["future_24"] / cagru_df["base_19"]) ** (1/4) - 1, 2),
        np.nan
    )

    agg_df["CAGR 5Y, руб"] = agg_df["CAGR 5Y, руб"].astype(float).map('{:.1%}'.format)
    agg_df["CAGR 5Y, уп"] = agg_df["CAGR 5Y, уп"].astype(float).map('{:.1%}'.format)
    
    st.subheader("Рыночное окружение")
    st.dataframe(agg_df)
 
    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
        # Write each dataframe to a different worksheet.
        agg_df.to_excel(writer, sheet_name='Sheet1', index=False)
        writer.close()

        download = st.download_button(
            label="Download as Excel",
            data=buffer,
            file_name='Сводная таблица.xlsx',
            #mime='application/vnd.ms-excel'
        )
    
else:
    st.warning("Выбери МНН или классификацию")
