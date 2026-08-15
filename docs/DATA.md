# DATA.md: Схема и контракты данных

## 1. Входные данные (Лист 'Data' в Excel)
| Колонка | Имя поля | Тип | Описание |
| :--- | :--- | :--- | :--- |
| A | id | String | Номер обращения |
| B | reg | DateTime | Дата регистрации |
| C | cls | DateTime | Дата закрытия |
| D | cli | String | Инициатор |
| E | anl | String | Аналитик |
| F | sec | String | Тема / Раздел |
| G | desc | String | Описание |
| H | sol | String | Решение |

## 2. Листы книги Excel
* **Logs:** Run_ID, Timestamp, Year, Week, Total_Tickets, Status.
* **Token:** Run_ID, Batch_Model, Batch_Tokens, Summary_Model, Summary_Tokens.