let
    подключение = Excel.Workbook(File.Contents(pSource09), null, true),
    сырые_данные = подключение{[Item="Сырье_заводы",Kind="Table"]}[Data],
    меняем_типы = Table.TransformColumnTypes(сырые_данные,{{"Федеральный округ", type text}, {"Показатель", type text}, {"Регион", type text}, {"Завод", type text}, {"Дата", type text}, {"Значение", type number}, {"Код сырья", type text}}),
    замена_значений = Table.ReplaceValue(меняем_типы,"Потребление в переработку","Потребление на переработку",Replacer.ReplaceText,{"Показатель"}),
    внешний_ключ = Table.CombineColumns(замена_значений, {"Регион", "Код сырья", "Дата"}, Combiner.CombineTextByDelimiter("_", QuoteStyle.None), "Ключ региона"),
    пивот_таблицы = Table.Pivot(внешний_ключ, List.Distinct(внешний_ключ[Показатель]), "Показатель", "Значение"),
    выбираем_столбцы = Table.SelectColumns(пивот_таблицы,{"Потребление на переработку", "Ключ региона", "Завод"}),
    заполняем_пустоты = Table.ReplaceValue(выбираем_столбцы,null,0,Replacer.ReplaceValue,{"Потребление на переработку"})
in
    заполняем_пустоты
