import argparse
from datetime import datetime
from pathlib import Path

import pandas as pd
from pptx import Presentation
from pptx.chart.data import CategoryChartData


def parse_args():
    parser = argparse.ArgumentParser(
        description="Расчёт Reach и TRP по целевой аудитории с формированием презентации"
    )
    parser.add_argument("target_audience", help="ЦА, например M_60+_A или W_25_45_BC")
    parser.add_argument("start_month", help="Начало периода в формате MM.YYYY")
    parser.add_argument("end_month", help="Конец периода в формате MM.YYYY")
    parser.add_argument("output_path", help="Путь для итоговой презентации")
    return parser.parse_args()


def parse_month(value):
    return pd.Timestamp(datetime.strptime(value, "%m.%Y")).normalize()


def parse_target_expression(expression):
    parts = expression.upper().split("_")
    gender = parts[0]
    age_expression = parts[1]
    income_group = None

    if age_expression.endswith("+"):
        age_from = int(age_expression[:-1])
        age_to = None
        if len(parts) == 3:
            income_group = parts[2]
    else:
        age_from = int(age_expression)
        age_to = int(parts[2])
        if len(parts) == 4:
            income_group = parts[3]

    return {
        "gender": gender,
        "age_from": age_from,
        "age_to": age_to,
        "income_group": income_group,
    }


def calculate_age(birth_dates, reference_date):
    age = reference_date.year - birth_dates.dt.year
    birthday_not_reached = (birth_dates.dt.month > reference_date.month) | (
        (birth_dates.dt.month == reference_date.month)
        & (birth_dates.dt.day > reference_date.day)
    )
    return age - birthday_not_reached.astype(int)


def describe_target(target):
    gender_names = {
        "M": "мужчины",
        "W": "женщины",
        "ALL": "все респонденты",
    }
    income_names = {
        "A": "доход группы A",
        "B": "доход группы B",
        "C": "доход группы C",
        "AB": "доход групп A или B",
        "BC": "доход групп B или C",
    }

    if target["age_to"] is None:
        age_text = f"от {target['age_from']} лет"
    else:
        age_text = f"от {target['age_from']} до {target['age_to']} лет"

    result = f"{gender_names[target['gender']]}, {age_text}"
    if target["income_group"] is not None:
        result += f", {income_names[target['income_group']]}"
    return result


def main():
    args = parse_args()

    target = parse_target_expression(args.target_audience)
    start_month = parse_month(args.start_month)
    end_month = parse_month(args.end_month)

    if start_month > end_month:
        raise SystemExit("Ошибка: начало периода не может быть позже конца периода")

    end_exclusive = end_month + pd.DateOffset(months=1)
    reference_date = end_exclusive - pd.Timedelta(days=1)

    # загрузка и подготовка исходных данных
    respondents = pd.read_csv("data/respondents.csv")
    views = pd.read_csv("data/views.csv")

    respondents["birth_date"] = pd.to_datetime(
        respondents["birth_date"], format="%d.%m.%Y"
    )
    views["date"] = pd.to_datetime(views["date"], format="%d.%m.%Y")
    respondents["age"] = calculate_age(respondents["birth_date"], reference_date)

    ca_mask = pd.Series(True, index=respondents.index)
    if target["gender"] != "ALL":
        ca_mask &= respondents["gender"] == target["gender"]
    ca_mask &= respondents["age"] >= target["age_from"]
    if target["age_to"] is not None:
        ca_mask &= respondents["age"] <= target["age_to"]
    if target["income_group"] is not None:
        ca_mask &= respondents["income_group"].isin(list(target["income_group"]))

    ca = respondents.loc[ca_mask].copy()
    if ca.empty:
        raise SystemExit("Ошибка: для заданных условий целевая аудитория пуста")

    views_in_period = views.loc[
        (views["date"] >= start_month) & (views["date"] < end_exclusive)
    ]
    ads_in_period = views_in_period[["advertiser", "ad_id"]].drop_duplicates()
    target_views = views_in_period.loc[
        views_in_period["respondent_id"].isin(ca["id"])
    ]

    # расчёт Reach по роликам и выбор пяти лучших рекламодателей
    unique_viewers_by_ad = (
        target_views.groupby(["advertiser", "ad_id"], as_index=False)["respondent_id"]
        .nunique()
        .rename(columns={"respondent_id": "unique_viewers"})
    )
    ads_with_viewers = ads_in_period.merge(
        unique_viewers_by_ad,
        on=["advertiser", "ad_id"],
        how="left",
    )
    ads_with_viewers["unique_viewers"] = ads_with_viewers["unique_viewers"].fillna(0)
    ads_with_viewers["reach"] = ads_with_viewers["unique_viewers"] / len(ca) * 100

    average_reach = (
        ads_with_viewers.groupby("advertiser", as_index=False)["reach"]
        .mean()
        .rename(columns={"reach": "average_reach"})
    )
    top5 = (
        average_reach.sort_values("average_reach", ascending=False)
        .head(5)
        .reset_index(drop=True)
    )

    # расчёт TRP по месяцам для пяти лучших рекламодателей
    top5_names = top5["advertiser"].tolist()
    top5_views = target_views.loc[
        target_views["advertiser"].isin(top5_names)
    ].copy()
    top5_views["month"] = top5_views["date"].dt.to_period("M")

    monthly_views = (
        top5_views.groupby(["advertiser", "month"])
        .size()
        .reset_index(name="view_count")
    )
    month_periods = pd.period_range(
        start=start_month.to_period("M"),
        end=end_month.to_period("M"),
        freq="M",
    )
    full_index = pd.MultiIndex.from_product(
        [top5_names, month_periods], names=["advertiser", "month"]
    )
    monthly_views = (
        monthly_views.set_index(["advertiser", "month"])
        .reindex(full_index, fill_value=0)
        .reset_index()
    )
    monthly_views["trp"] = monthly_views["view_count"] / len(ca) * 100
    trp_table = (
        monthly_views.pivot(
            index="month", columns="advertiser", values="trp"
        )
        .reindex(index=month_periods, columns=top5_names, fill_value=0)
        .fillna(0)
    )

    # заполнение шаблона презентации
    presentation = Presentation("template/presentation_template.pptx")
    subtitle = presentation.slides[0].shapes[1].text_frame
    subtitle.text = (
        f"Целевая аудитория: {describe_target(target)}\n"
        f"Период: {start_month.strftime('%m.%Y')} – {end_month.strftime('%m.%Y')}"
    )

    table = presentation.slides[1].shapes[1].table
    for index, row in top5.iterrows():
        table.cell(index + 1, 0).text = str(row["advertiser"])
        table.cell(index + 1, 1).text = f"{row['average_reach']:.2f}%"

    chart = presentation.slides[2].shapes[1].chart
    chart_data = CategoryChartData()
    chart_data.categories = [
        period.strftime("%m.%Y") for period in trp_table.index
    ]
    for advertiser in top5_names:
        chart_data.add_series(advertiser, trp_table[advertiser].tolist())
    chart.replace_data(chart_data)
    chart.plots[0].has_data_labels = False

    output_path = Path(args.output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    presentation.save(output_path)

    print(f"ЦА: {len(ca)} респондентов")
    print(top5.to_string(index=False))
    print(f"Презентация сохранена: {output_path}")


if __name__ == "__main__":
    main()
