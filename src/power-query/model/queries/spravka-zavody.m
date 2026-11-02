let
    подключение = Excel.Workbook(File.Contents(pSource07), null, true),
    получение_данных = подключение{[Item="Таблица1",Kind="Table"]}[Data],
    меняем_типы = Table.TransformColumnTypes(получение_данных,{{"Регион", type text}, {"Холдинг", type text}, {"Завод", type text}}),
    выбираем_столбцы = Table.SelectColumns(меняем_типы,{"Холдинг", "Завод"})
in
    выбираем_столбцы
