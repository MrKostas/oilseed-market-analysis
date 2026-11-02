let
    подключение = List.Distinct(Масло_регионы[#"Код сырья"]),
    список_таблица = Table.FromList(подключение),
    столбец_сырье = Table.AddColumn(список_таблица, "Сырье", each if [Column1] = "п" then "Подсолнечник"
                                    else if [Column1] = "с" then "Соя" else "Рапс"),
    меняем_типы_названия = Table.RenameColumns(Table.TransformColumnTypes(столбец_сырье,{{"Column1", type text}, {"Сырье", type text}}),{{"Column1", "Код сырья"}})
in
    меняем_типы_названия
