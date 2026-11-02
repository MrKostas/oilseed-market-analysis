let
    скученные_данные = Excel.Workbook(File.Contents(pSource02), null, true){[Name = "Свод рапс"]}[Data],
    сырые_данные = Table.Skip(скученные_данные, List.PositionOf(скученные_данные[Column1], "Баланс рапса, тысяч тонн")),
    столбец_сырье = Table.FillDown(Table.AddColumn(сырые_данные, "Column" & Number.ToText(Number.FromText(Text.AfterDelimiter(List.Last(Table.ColumnNames(сырые_данные)), "n"))+1), each if Text.BeforeDelimiter([Column1], " ") = "Баланс" then Text.Start(Text.BetweenDelimiters([Column1], " ", " "), 1) else null), {"Column" & Number.ToText(Number.FromText(Text.AfterDelimiter(List.Last(Table.ColumnNames(сырые_данные)), "n"))+1)}),
    поднимаем_заголовки = Table.PromoteHeaders(столбец_сырье, [PromoteAllScalars = true]),
    фильтрация_строк_1 = Table.SelectRows(поднимаем_заголовки, each [#"Баланс рапса, тысяч тонн"] <> ""),
    список_заголовков = List.Select(Table.ColumnNames(фильтрация_строк_1), each Text.Start(_, 2) = "Ба" or Text.Start(_, 2) = "01" or Text.Start(_, 1) = "р"),
    выбираем_столбцы = Table.SelectColumns(фильтрация_строк_1, список_заголовков),
    меняем_название_столбца = Table.RenameColumns(выбираем_столбцы, {"р", "Код сырья"}),
    столбец_фо = Table.AddColumn(меняем_название_столбца, "Федеральный округ", each if Text.AfterDelimiter([#"Баланс рапса, тысяч тонн"], " ") = "федеральный округ" then [#"Баланс рапса, тысяч тонн"] else null),
    столбец_показатель = Table.AddColumn(столбец_фо, "Показатель", each if List.Contains({"Импорт", "Экспорт", "Запасы на начало месяца",
                                                                                          "Валовый сбор", "Потребление", "Остатки на конец месяца",
                                                                                          "Приток/отток (РФ)", "Потребление на переработку", "Потребление на семена",
                                                                                          "Прочее потребление"}, [#"Баланс рапса, тысяч тонн"]) then [#"Баланс рапса, тысяч тонн"] else null),
    заполнение_показатель_фо = Table.FillDown(столбец_показатель, {"Федеральный округ", "Показатель"}),
    фильтрация_строк_2 = Table.SelectRows(заполнение_показатель_фо, each Text.AfterDelimiter([#"Баланс рапса, тысяч тонн"], " ") <> "федеральный округ" and [#"Баланс рапса, тысяч тонн"] <> null and Text.BeforeDelimiter([#"Баланс рапса, тысяч тонн"], " ") <> "Баланс"
                                                                  and not List.Contains({"Импорт", "Экспорт", "Запасы на начало месяца",
                                                                                          "Валовый сбор", "Потребление", "Остатки на конец месяца",
                                                                                          "Приток/отток (РФ)", "Потребление на переработку", "Потребление на семена",
                                                                                          "Прочее потребление"}, [#"Баланс рапса, тысяч тонн"])),
    столбец_регион = Table.FillDown(Table.AddColumn(фильтрация_строк_2, "Регион", each if List.Contains({"область", "республика", "автономный", "авт.округ", "край"}, Text.BetweenDelimiters([#"Баланс рапса, тысяч тонн"], " ", " "), Comparer.OrdinalIgnoreCase) and Text.BeforeDelimiter([#"Баланс рапса, тысяч тонн"], " ") <> "Прочие" then [#"Баланс рапса, тысяч тонн"] else if List.Contains({"область", "республика", "автономный", "авт.округ", "край"}, Text.BeforeDelimiter([#"Баланс рапса, тысяч тонн"], " "), Comparer.OrdinalIgnoreCase) and Text.BeforeDelimiter([#"Баланс рапса, тысяч тонн"], " ") <> "Прочие" then [#"Баланс рапса, тысяч тонн"] else if List.Contains({"область", "республика", "автономный", "авт.округ", "край"}, Text.AfterDelimiter([#"Баланс рапса, тысяч тонн"], " ", 1), Comparer.OrdinalIgnoreCase) and Text.BeforeDelimiter([#"Баланс рапса, тысяч тонн"], " ") <> "Прочие" then [#"Баланс рапса, тысяч тонн"] else if List.Contains({"область", "республика", "автономный", "авт.округ", "край"}, Text.BetweenDelimiters([#"Баланс рапса, тысяч тонн"], " ", "-"), Comparer.OrdinalIgnoreCase) and Text.BeforeDelimiter([#"Баланс рапса, тысяч тонн"], " ") <> "Прочие" then [#"Баланс рапса, тысяч тонн"] else null), {"Регион"}),
    столбец_завод = Table.AddColumn(столбец_регион, "Завод", each if [#"Баланс рапса, тысяч тонн"] <> [Регион] then [#"Баланс рапса, тысяч тонн"] else null),
    фильтрация_строк_3 = Table.SelectRows(столбец_завод, each [Завод] = null),
    выбираем_столбцы_2 = Table.SelectColumns(фильтрация_строк_3, List.Select(Table.ColumnNames(фильтрация_строк_3), each _ <> "Завод" and _ <> "Баланс рапса, тысяч тонн")),
    анпивот_таблицы = Table.Unpivot(выбираем_столбцы_2, List.Select(Table.ColumnNames(выбираем_столбцы_2), each Text.Start(_, 2) = "01"), "Дата", "Значение"),
    меняем_типы = Table.TransformColumnTypes(анпивот_таблицы, {{"Код сырья", type text}, {"Федеральный округ", type text},
                                                               {"Показатель", type text}, {"Регион", type text}, {"Дата", type date}, {"Значение", type number}})
in
    меняем_типы
