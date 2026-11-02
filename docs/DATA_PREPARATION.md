# Подготовка данных в Power Query M

## Поток

Исходные Excel-балансы → восстановление заголовков и групп, FillDown, Unpivot → отдельные региональные/заводские таблицы → промежуточные книги → Pivot показателей, составные ключи и справочники в Power BI → DAX → три аналитические страницы.

## Два этапа обработки

- `source` — выражения из DataMashup внутри подготовительных Excel-книг. Это обработка первичных балансов/мониторинга, а не строки данных из листов.
- `model` — выражения Power Query семантической модели PBIX. Они читают промежуточные книги и формируют факты/справочники отчёта.
- `parameters` — добавленные текстовые параметры путей. Имена исходных запросов и преобразования сохранены; литералы локальных/сетевых путей заменены параметрами. Иного исправления логики не выполнялось.

Промежуточные XLSX могут содержать сохранённые значения. Они не входят в repository. У части справочников нет DataMashup: отсутствие запроса не означает пропуск извлечения.

## Что делает подготовка

- Таблицы масел сначала разделяются на региональные и заводские показатели. Список отчётных столбцов определяется по заголовкам, даты разворачиваются в длинный формат.
- Подсолнечник, соя и рапс обрабатываются отдельными запросами и объединяются в Сырье_регионы / Сырье_заводы.
- Для разных видов шрота есть отдельные региональные/заводские потоки и запросы потребления; их имена не следует смешивать с именами таблиц конечной модели.
- Диагностический запрос ошибок типов сохранён как отдельный артефакт. Он не является дополнительной таблицей аналитической модели.

## Параметры источников

Значение каждого параметра при поставке — `REPLACE_WITH_LOCAL_SOURCE_PATH`. Заменить его локальным путём в своей разрешённой среде; реальные пути в код не включены. Один исходный путь соответствует одному параметру внутри этого кейса. Этапы Excel и Power BI используют отдельные пространства запросов.

| Параметр | Где используется |
|---|---|
| `pSource01` | source: Масло_регионы; source: Масло_заводы |
| `pSource02` | source: Подсолнечник_регионы; source: Подсолнечник_заводы; source: Соя_регионы; source: Соя_заводы; source: Рапс_регионы; source: Рапс_заводы |
| `pSource03` | source: СШ_регионы; source: СШ_заводы; source: Потребление_СШ |
| `pSource04` | source: РШ_регионы; source: РШ_заводы; source: Потребление_РШ |
| `pSource05` | source: ПШ_регионы; source: ПШ_заводы; source: Потребление_ПШ |
| `pSource06` | model: Масло_регионы; model: Масло_заводы |
| `pSource07` | model: Справка_заводы; model: Справка_регионы |
| `pSource08` | model: Шроты_заводы; model: Шроты_регионы |
| `pSource09` | model: Сырье_заводы; model: Сырье_регионы |

## Реестр запросов

| Этап | Имя запроса в Power Query | Назначение | Файл |
|---|---|---|---|
| source | `Масло_регионы` | Преобразование | [M-код](../src/power-query/source/dataset-masla/maslo-regiony.m) |
| source | `Масло_заводы` | Преобразование | [M-код](../src/power-query/source/dataset-masla/maslo-zavody.m) |
| source | `Ошибки в Масло_регионы` | Диагностика | [M-код](../src/power-query/source/dataset-masla/oshibki-v-maslo-regiony.m) |
| source | `Подсолнечник_регионы` | Преобразование | [M-код](../src/power-query/source/dataset-maslichnye/podsolnechnik-regiony.m) |
| source | `Подсолнечник_заводы` | Преобразование | [M-код](../src/power-query/source/dataset-maslichnye/podsolnechnik-zavody.m) |
| source | `Соя_регионы` | Преобразование | [M-код](../src/power-query/source/dataset-maslichnye/soya-regiony.m) |
| source | `Соя_заводы` | Преобразование | [M-код](../src/power-query/source/dataset-maslichnye/soya-zavody.m) |
| source | `Рапс_регионы` | Преобразование | [M-код](../src/power-query/source/dataset-maslichnye/raps-regiony.m) |
| source | `Рапс_заводы` | Преобразование | [M-код](../src/power-query/source/dataset-maslichnye/raps-zavody.m) |
| source | `Сырье_регионы` | Преобразование | [M-код](../src/power-query/source/dataset-maslichnye/syr-e-regiony.m) |
| source | `Сырье_заводы` | Преобразование | [M-код](../src/power-query/source/dataset-maslichnye/syr-e-zavody.m) |
| source | `СШ_регионы` | Преобразование | [M-код](../src/power-query/source/dataset-shrota/ssh-regiony.m) |
| source | `СШ_заводы` | Преобразование | [M-код](../src/power-query/source/dataset-shrota/ssh-zavody.m) |
| source | `РШ_регионы` | Преобразование | [M-код](../src/power-query/source/dataset-shrota/rsh-regiony.m) |
| source | `РШ_заводы` | Преобразование | [M-код](../src/power-query/source/dataset-shrota/rsh-zavody.m) |
| source | `ПШ_регионы` | Преобразование | [M-код](../src/power-query/source/dataset-shrota/psh-regiony.m) |
| source | `ПШ_заводы` | Преобразование | [M-код](../src/power-query/source/dataset-shrota/psh-zavody.m) |
| source | `Шрота_регионы` | Преобразование | [M-код](../src/power-query/source/dataset-shrota/shrota-regiony.m) |
| source | `Шрота_заводы` | Преобразование | [M-код](../src/power-query/source/dataset-shrota/shrota-zavody.m) |
| source | `Потребление_СШ` | Преобразование | [M-код](../src/power-query/source/dataset-shrota/potreblenie-ssh.m) |
| source | `Потребление_ПШ` | Преобразование | [M-код](../src/power-query/source/dataset-shrota/potreblenie-psh.m) |
| source | `Потребление_РШ` | Преобразование | [M-код](../src/power-query/source/dataset-shrota/potreblenie-rsh.m) |
| model | `Масло_регионы` | Преобразование | [M-код](../src/power-query/model/queries/maslo-regiony.m) |
| model | `Масло_заводы` | Преобразование | [M-код](../src/power-query/model/queries/maslo-zavody.m) |
| model | `Справка_дата` | Преобразование | [M-код](../src/power-query/model/queries/spravka-data.m) |
| model | `Справка_заводы` | Преобразование | [M-код](../src/power-query/model/queries/spravka-zavody.m) |
| model | `Справка_регионы` | Преобразование | [M-код](../src/power-query/model/queries/spravka-regiony.m) |
| model | `Справка_сырье` | Преобразование | [M-код](../src/power-query/model/queries/spravka-syr-e.m) |
| model | `Шроты_заводы` | Преобразование | [M-код](../src/power-query/model/queries/shroty-zavody.m) |
| model | `Шроты_регионы` | Преобразование | [M-код](../src/power-query/model/queries/shroty-regiony.m) |
| model | `Сырье_заводы` | Преобразование | [M-код](../src/power-query/model/queries/syr-e-zavody.m) |
| model | `Сырье_регионы` | Преобразование | [M-код](../src/power-query/model/queries/syr-e-regiony.m) |

Полный машиночитаемый [query-index.csv](../src/power-query/query-index.csv) сохраняет связь имени запроса и файла. Диагностические запросы не надо загружать как факты модели.

[Порядок восстановления и ограничения](REPRODUCING.md) · [Обзор](../README.md).
