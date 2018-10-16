# Excel-SQL
Function to perform SQL queries within Excel, returning result to range or as array

## Parameters
ExcelSQL(query, result_location, header, as_array)

| Name              | Required/Optional   | Data Type   | Description                                                                     |
| :---              | :---                | :---        | :---                                                                            |
| query             | Required            | String      | String of SQL query; uses MS Access dialect.                                    |
| result_location   | Optional            | Range       | Top left corner of where query result will be outputted to worksheet. Used only                                                             when as_array = False.                                                          |
| header            | Optional            | Boolean     | Include column labels in query result. True by default.                         |
| as_array          | Optional            | Boolean     | Return query result as Variant for further use in VBA script. False by default. |
