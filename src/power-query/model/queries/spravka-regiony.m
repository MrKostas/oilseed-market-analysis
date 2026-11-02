let
    Источник = Excel.Workbook(File.Contents(pSource07), null, true),
    Таблица2_Table = Источник{[Item="Таблица2",Kind="Table"]}[Data],
    #"Измененный тип" = Table.TransformColumnTypes(Таблица2_Table,{{"Регион", type text}, {"Округ", type text}, {"Округ_краткий", type text}, {"Страна", type text}})
in
    #"Измененный тип"
