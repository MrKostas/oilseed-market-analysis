let
    подключение = Excel.Workbook(File.Contents(pSource06), null, true),
    получение_данных = подключение{[Item="Масло_заводы",Kind="Table"]}[Data],
    измененные_типы = Table.TransformColumnTypes(получение_данных,{{"Код сырья", type text}, {"Федеральный округ", type text}, {"Показатель", type text}, {"Регион", type text}, {"Завод", type text}, {"Дата", type text}, {"Значение", type number}}),
    внешний_ключ = Table.CombineColumns(измененные_типы, {"Регион", "Код сырья", "Дата"}, Combiner.CombineTextByDelimiter("_", QuoteStyle.None), "Ключ региона"),
    пивот_таблицы = Table.Pivot(внешний_ключ, List.Distinct(внешний_ключ[Показатель]), "Показатель", "Значение")
in
    пивот_таблицы
