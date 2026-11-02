let
    подключение = Excel.Workbook(File.Contents(pSource08), null, true),
    получение_данных = подключение{[Item="Шрота_заводы_2",Kind="Table"]}[Data],
    измененные_типы = Table.TransformColumnTypes(получение_данных,{{"Федеральный округ", type text}, {"Показатель", type text}, {"Регион", type text}, {"Завод", type text}, {"Дата", type text}, {"Значение", type number}, {"Код сырья", type text}}),
    внешний_ключ = Table.CombineColumns(измененные_типы, {"Регион", "Код сырья", "Дата"}, Combiner.CombineTextByDelimiter("_", QuoteStyle.None), "Ключ региона"),
    пивот_таблицы = Table.Pivot(внешний_ключ, List.Distinct(внешний_ключ[Показатель]), "Показатель", "Значение"),
    заполнение_пустот = Table.ReplaceValue(пивот_таблицы,null,0,Replacer.ReplaceValue,{"Экспорт"}),
    #"Другие удаленные столбцы" = Table.SelectColumns(заполнение_пустот,{"Завод", "Ключ региона", "Производство", "Экспорт"}),
    #"Переименованные столбцы" = Table.RenameColumns(#"Другие удаленные столбцы",{{"Экспорт", "Экспорт_завод"}})
in
    #"Переименованные столбцы"
