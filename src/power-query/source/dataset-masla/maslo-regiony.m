let
    скученные_данные = Table.SelectColumns(Table.SelectRows(Excel.Workbook(File.Contents(pSource01), null, true), each Text.Start([Name], 4) = "Свод"), {"Data"}),
    сырые_данные = Table.Skip(Table.ExpandTableColumn(скученные_данные, "Data", Table.ColumnNames(скученные_данные{0}[Data])), 2),
    столбец_сырье = Table.FillDown(Table.AddColumn(сырые_данные, "Column" & Number.ToText(Number.FromText(Text.AfterDelimiter(List.Last(Table.ColumnNames(сырые_данные)), "n"))+1), each if Text.BeforeDelimiter([Column1], " ") = "Баланс" then Text.Start(Text.BetweenDelimiters([Column1], " ", " "), 1) else null), {"Column" & Number.ToText(Number.FromText(Text.AfterDelimiter(List.Last(Table.ColumnNames(сырые_данные)), "n"))+1)}),
    поднимаем_заголовки = Table.PromoteHeaders(столбец_сырье, [PromoteAllScalars = true]),
    фильтрация_строк_1 = Table.SelectRows(поднимаем_заголовки, each [#"Баланс подсолнечного масла, тысяч тонн"] <> ""),
    список_заголовков = List.Select(Table.ColumnNames(фильтрация_строк_1), each Text.Start(_, 2) = "Ба" or Text.Start(_, 2) = "01" or Text.Start(_, 1) = "п"),
    выбираем_столбцы = Table.SelectColumns(фильтрация_строк_1, список_заголовков),
    меняем_название_столбца = Table.RenameColumns(выбираем_столбцы, {"п", "Код сырья"}),
    столбец_фо = Table.AddColumn(меняем_название_столбца, "Федеральный округ", each if Text.AfterDelimiter([#"Баланс подсолнечного масла, тысяч тонн"], " ") = "федеральный округ" then [#"Баланс подсолнечного масла, тысяч тонн"] else null),
    столбец_показатель = Table.AddColumn(столбец_фо, "Показатель", each if List.Contains({"Импорт", "Экспорт", "Запасы на начало месяца",
                                                                                          "Производство", "Потребление", "Остатки на конец месяца",
                                                                                          "Приток/отток (РФ)", "Потребление на промышленные жиры", "Потребление на фасовку",
                                                                                          "Потребление на майонез", "Потребление на комбикорма", "Прочее потребление"}, [#"Баланс подсолнечного масла, тысяч тонн"]) then [#"Баланс подсолнечного масла, тысяч тонн"] else null),
    заполнение_показатель_фо = Table.FillDown(столбец_показатель, {"Федеральный округ", "Показатель"}),
    фильтрация_строк_2 = Table.SelectRows(заполнение_показатель_фо, each Text.AfterDelimiter([#"Баланс подсолнечного масла, тысяч тонн"], " ") <> "федеральный округ" and [#"Баланс подсолнечного масла, тысяч тонн"] <> null and Text.BeforeDelimiter([#"Баланс подсолнечного масла, тысяч тонн"], " ") <> "Баланс"
                                                                  and not List.Contains({"Импорт", "Экспорт", "Запасы на начало месяца",
                                                                                          "Производство", "Потребление", "Остатки на конец месяца",
                                                                                          "Приток/отток (РФ)", "Потребление на промышленные жиры", "Потребление на фасовку",
                                                                                          "Потребление на майонез", "Потребление на комбикорма", "Прочее потребление"}, [#"Баланс подсолнечного масла, тысяч тонн"])),
    столбец_регион = Table.FillDown(Table.AddColumn(фильтрация_строк_2, "Регион", each if List.Contains({"область", "республика", "автономный", "авт.округ", "край"}, Text.BetweenDelimiters([#"Баланс подсолнечного масла, тысяч тонн"], " ", " "), Comparer.OrdinalIgnoreCase) and Text.BeforeDelimiter([#"Баланс подсолнечного масла, тысяч тонн"], " ") <> "Прочие" then [#"Баланс подсолнечного масла, тысяч тонн"] else if List.Contains({"область", "республика", "автономный", "авт.округ", "край"}, Text.BeforeDelimiter([#"Баланс подсолнечного масла, тысяч тонн"], " "), Comparer.OrdinalIgnoreCase) and Text.BeforeDelimiter([#"Баланс подсолнечного масла, тысяч тонн"], " ") <> "Прочие" then [#"Баланс подсолнечного масла, тысяч тонн"] else if List.Contains({"область", "республика", "автономный", "авт.округ", "край"}, Text.AfterDelimiter([#"Баланс подсолнечного масла, тысяч тонн"], " ", 1), Comparer.OrdinalIgnoreCase) and Text.BeforeDelimiter([#"Баланс подсолнечного масла, тысяч тонн"], " ") <> "Прочие" then [#"Баланс подсолнечного масла, тысяч тонн"] else if List.Contains({"область", "республика", "автономный", "авт.округ", "край"}, Text.BetweenDelimiters([#"Баланс подсолнечного масла, тысяч тонн"], " ", "-"), Comparer.OrdinalIgnoreCase) and Text.BeforeDelimiter([#"Баланс подсолнечного масла, тысяч тонн"], " ") <> "Прочие" then [#"Баланс подсолнечного масла, тысяч тонн"] else null), {"Регион"}),
    столбец_завод = Table.AddColumn(столбец_регион, "Завод", each if [#"Баланс подсолнечного масла, тысяч тонн"] <> [Регион] then [#"Баланс подсолнечного масла, тысяч тонн"] else null),
    фильтрация_строк_3 = Table.SelectRows(столбец_завод, each [Завод] = null),
    выбираем_столбцы_2 = Table.SelectColumns(фильтрация_строк_3, List.Select(Table.ColumnNames(фильтрация_строк_3), each _ <> "Завод" and _ <> "Баланс подсолнечного масла, тысяч тонн")),
    анпивот_таблицы = Table.Unpivot(выбираем_столбцы_2, List.Select(Table.ColumnNames(выбираем_столбцы_2), each Text.Start(_, 2) = "01"), "Дата", "Значение"),
    меняем_типы = Table.TransformColumnTypes(анпивот_таблицы, {{"Код сырья", type text}, {"Федеральный округ", type text},
                                                               {"Показатель", type text}, {"Регион", type text}, {"Дата", type date}, {"Значение", type number}})
in
    меняем_типы
