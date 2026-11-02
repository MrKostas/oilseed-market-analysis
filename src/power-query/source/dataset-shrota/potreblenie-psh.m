let
    подключение = Excel.Workbook(File.Contents(pSource05), null, true){[Name = "Потребление расчет"]}[Data],
    удаление_строк = Table.SelectRows(Table.Skip(подключение, List.PositionOf(подключение[Column1], "Свинина_1")), each [Column1] <> null),
    столбец_животное = Table.FillDown(Table.AddColumn(удаление_строк, "Животное", each if List.Contains({"Свинина_1", "Яйца", "Говядина", "Птица", "Молоко", "Рыба (карповые)", "Рыба (лососевые)"}, [Column1]) then [Column1] else null), {"Животное"}),
    поднимаем_заголовки = Table.PromoteHeaders(столбец_животное, [PromoteAllScalars = true]),
    переименовываем_столбцы = Table.RenameColumns(поднимаем_заголовки,{{"Свинина_1_1", "Показатель"}}),
    список_заголовков = List.Select(Table.ColumnNames(переименовываем_столбцы), each Text.Start(_, 6) <> "Column" and Text.Start(_, 1) <> "2"),
    выбираем_столбцы = Table.SelectColumns(переименовываем_столбцы, список_заголовков),
    столбец_фо = Table.AddColumn(выбираем_столбцы, "Федеральный округ", each if Text.AfterDelimiter([#"Свинина_1"], " ") = "федеральный округ" then [#"Свинина_1"] else null),
    заполнение_фо = Table.FillDown(столбец_фо, {"Федеральный округ"}),
    столбец_регион = Table.FillDown(Table.AddColumn(заполнение_фо, "Регион", each if List.Contains({"область", "республика", "автономный", "авт.округ", "край"}, Text.BetweenDelimiters([#"Свинина_1"], " ", " "), Comparer.OrdinalIgnoreCase) then [#"Свинина_1"] else if List.Contains({"область", "республика", "автономный", "авт.округ", "край"}, Text.BeforeDelimiter([#"Свинина_1"], " "), Comparer.OrdinalIgnoreCase) then [#"Свинина_1"] else if List.Contains({"область", "республика", "автономный", "авт.округ", "край"}, Text.AfterDelimiter([#"Свинина_1"], " ", 1), Comparer.OrdinalIgnoreCase) then [#"Свинина_1"] else if List.Contains({"область", "республика", "автономный", "авт.округ", "край"}, Text.BetweenDelimiters([#"Свинина_1"], " ", "-"), Comparer.OrdinalIgnoreCase) then [#"Свинина_1"] else null), {"Регион"}),
    фильтрация_строк = Table.SelectRows(столбец_регион, each [Свинина_1] <> [#"Федеральный округ"]),
    фильтрация_строк_2 = Table.SelectRows(фильтрация_строк, each ([Федеральный округ] <> null) and ([Свинина_1] <> "" and [Свинина_1] <> "Говядина" and [Свинина_1] <> "Молоко" and [Свинина_1] <> "Птица" and [Свинина_1] <> "Российская Федерация" and [Свинина_1] <> "Рыба (карповые)" and [Свинина_1] <> "Рыба (лососевые)" and [Свинина_1] <> "Яйца")),
    удаляем_столбец = Table.RemoveColumns(фильтрация_строк_2,{"Свинина_1"}),
    анпивот_таблицы = Table.UnpivotOtherColumns(удаляем_столбец, {"Показатель", "Федеральный округ", "Регион"}, "Дата", "Значение"),
    столбец_сырья = Table.AddColumn(анпивот_таблицы, "Код сырья", each "п"),
    меняем_места = Table.ReorderColumns(столбец_сырья,{"Федеральный округ", "Показатель", "Регион", "Дата", "Значение", "Код сырья"}),
    меняем_типы = Table.TransformColumnTypes(меняем_места,{{"Код сырья", type text}, {"Регион", type text}, {"Федеральный округ", type text}, {"Показатель", type text}, {"Дата", type date}, {"Значение", type number}}),
    убираем_знаки = Table.ReplaceValue(меняем_типы, "_1", "", Replacer.ReplaceText, {"Показатель"})
in
    убираем_знаки
