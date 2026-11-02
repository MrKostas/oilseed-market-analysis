let
    список_дат = List.Distinct(Масло_регионы[Дата]),
    список_таблица = Table.TransformColumnTypes(Table.FromList(список_дат, Splitter.SplitByNothing(), null, null, ExtraValues.Error), {"Column1", type date}),
    столбец_год = Table.AddColumn(список_таблица, "Год", each Date.Year([Column1]), Int64.Type),
    столбец_месяц = Table.AddColumn(столбец_год, "Название месяца", each Date.MonthName([Column1]), type text),
    измененные_названия = Table.RenameColumns(столбец_месяц,{{"Column1", "Дата"}, {"Название месяца", "Месяц"}}),
    столбец_сезон = Table.AddColumn(измененные_названия, "Сезон", each if [Дата] >= #date(2020, 09, 01) and [Дата] < #date(2021, 09, 01) then "20/21"
                                                                  else if [Дата] >= #date(2021, 09, 01) and [Дата] < #date(2022, 09, 01) then "21/22"
                                                                  else if [Дата] >= #date(2022, 09, 01) and [Дата] < #date(2023, 09, 01) then "22/23"
                                                                  else if [Дата] >= #date(2023, 09, 01) and [Дата] < #date(2024, 09, 01) then "23/24"
                                                                  else if [Дата] >= #date(2024, 09, 01) and [Дата] < #date(2025, 09, 01) then "24/25"
                                                                  else if [Дата] >= #date(2025, 09, 01) and [Дата] < #date(2026, 09, 01) then "25/26"
                                                                  else if [Дата] >= #date(2026, 09, 01) and [Дата] < #date(2027, 09, 01) then "26/27"
                                                                  else if [Дата] >= #date(2027, 09, 01) and [Дата] < #date(2028, 09, 01) then "27/28"
                                                                  else "Добавьте новый сезон", type text)
in
    столбец_сезон
