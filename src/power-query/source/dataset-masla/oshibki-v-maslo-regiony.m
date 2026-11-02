let
Источник = Масло_регионы,
  #"Обнаруженные несоответствия типов" = let
    tableWithOnlyPrimitiveTypes = Table.SelectColumns(Источник, Table.ColumnsOfType(Источник, {type nullable number, type nullable text, type nullable logical, type nullable date, type nullable datetime, type nullable datetimezone, type nullable time, type nullable duration})),
    recordTypeFields = Type.RecordFields(Type.TableRow(Value.Type(tableWithOnlyPrimitiveTypes))),
    fieldNames = Record.FieldNames(recordTypeFields),
    fieldTypes = List.Transform(Record.ToList(recordTypeFields), each [Type]),
    pairs = List.Transform(List.Positions(fieldNames), (i) => {fieldNames{i}, (v) => if v = null or Value.Is(v, fieldTypes{i}) then v else error [Message = "Тип значения не соответствует типу столбца.", Detail = v], fieldTypes{i}})
in
    Table.TransformColumns(Источник, pairs),
  #"Добавлен индекс" = Table.AddIndexColumn(#"Обнаруженные несоответствия типов", "Номер строки" ,1),
  #"Сохраненные ошибки" = Table.SelectRowsWithErrors(#"Добавлен индекс", {"Код сырья", "Федеральный округ", "Показатель", "Регион", "Дата", "Значение"}),
  #"Переупорядоченные столбцы" = Table.ReorderColumns(#"Сохраненные ошибки", {"Номер строки", "Код сырья", "Федеральный округ", "Показатель", "Регион", "Дата", "Значение"})
in
  #"Переупорядоченные столбцы"
