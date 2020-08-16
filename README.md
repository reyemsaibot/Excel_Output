# Excel_Output
Output of an internal table in Microsoft Excel.

Just declare use the following code snippet 

```ABAP
Data: lo_excel type ref to zcl_excel.

lo_excel = NEW #( ).

lo_excel->output_to_excel(<br>
  EXPORTING<br>
    it_table           = lt_output "Internal Table<br>
    iv_sheet_name      = 'TEST'    "Sheet Name<br>
    iv_group_condition = 'ABC'     "Group results in different sheets (optional)<br>
).
```
