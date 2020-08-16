# Excel Output
Output of an internal table in Microsoft Excel.

Just declare use the following code snippet 

```ABAP
Data: lo_excel type ref to zcl_excel.

lo_excel = NEW #( ).

lo_excel->output_to_excel(
  EXPORTING
    it_table           = lt_output "Internal Table
    iv_sheet_name      = 'TEST'    "Sheet Name
    iv_group_condition = 'ABC'     "Group results in different sheets (optional)
).
```
