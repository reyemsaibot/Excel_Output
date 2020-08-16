"#autoformat
CLASS zcl_excel DEFINITION
  PUBLIC
  FINAL
  CREATE PUBLIC .

  PUBLIC SECTION.

    TYPES ty_excel_cell      TYPE c LENGTH 4096.
    TYPES tyt_excel_cell     TYPE TABLE OF ty_excel_cell.

    METHODS output_to_excel
      IMPORTING
        !it_table           TYPE ANY TABLE
        !iv_sheet_name      TYPE string
        !iv_group_condition TYPE string OPTIONAL .

  PROTECTED SECTION.
  PRIVATE SECTION.
    METHODS get_excel
      EXPORTING
        eo_sheets           TYPE ole2_object
        eo_excel            TYPE ole2_object
        ev_number_of_sheets TYPE i.

    METHODS get_active_sheet
      IMPORTING
        iv_sheet_index TYPE i
      CHANGING
        co_excel       TYPE ole2_object
        co_sheets      TYPE ole2_object
        co_sheet       TYPE ole2_object.

    METHODS set_header
      IMPORTING
        io_ref_struc TYPE REF TO cl_abap_structdescr
      CHANGING
        co_sheet     TYPE ole2_object.

    METHODS generate_output_table
      IMPORTING
        it_table        TYPE ANY TABLE

        io_ref_struc    TYPE REF TO cl_abap_structdescr
      EXPORTING
        et_output_table TYPE tyt_excel_cell.

    METHODS fill_sheet
      IMPORTING
        it_output_table TYPE tyt_excel_cell

        io_ref_struc    TYPE REF TO cl_abap_structdescr
        iv_sheet_name   TYPE string
      CHANGING
        co_sheet        TYPE ole2_object.

    METHODS get_number_of_columns
      IMPORTING
        io_ref_struc                TYPE REF TO cl_abap_structdescr
      RETURNING
        VALUE(rv_number_of_columns) TYPE i.

    METHODS adjust_column_width
      IMPORTING
        iv_column_index TYPE i
        iv_width        TYPE i
      CHANGING
        co_worksheet    TYPE ole2_object.

    METHODS autoformat_column_width
      IMPORTING
        iv_column_index TYPE i
      CHANGING
        co_sheet        TYPE ole2_object.

    METHODS return_subtable
      IMPORTING
        it_all_data  TYPE ANY TABLE
        it_key_field TYPE string
        it_key_value TYPE string
      EXPORTING
        et_subtable  TYPE ANY TABLE.

ENDCLASS.



CLASS zcl_excel IMPLEMENTATION.


  METHOD adjust_column_width.
    DATA: lo_column    TYPE ole2_object.

    "Select the Column
    CALL METHOD OF co_worksheet 'Columns'    = lo_column
                   EXPORTING    #1           = iv_column_index.

    CALL METHOD OF lo_column 'select'.

    SET PROPERTY OF lo_column 'ColumnWidth' = iv_width.

  ENDMETHOD.


  METHOD autoformat_column_width.
    DATA: lo_column      TYPE ole2_object.

    "Select the Column
    CALL METHOD OF
        co_sheet
        'Columns' = lo_column
      EXPORTING
        #1        = iv_column_index.

    CALL METHOD OF
      lo_column
      'Autofit'.
  ENDMETHOD.


  METHOD fill_sheet.
    DATA: lo_cell       TYPE ole2_object.
    DATA: lv_sheet_name TYPE string.

**********************************************************************
* Create Header
**********************************************************************

    set_header( EXPORTING io_ref_struc = io_ref_struc
                CHANGING  co_sheet     = co_sheet ).

**********************************************************************
* Add internal table (=paste content) into clipboard
**********************************************************************

    CALL FUNCTION 'CLPB_EXPORT'
      TABLES
        data_tab   = it_output_table
      EXCEPTIONS
        clpb_error = 1
        OTHERS     = 2.

**********************************************************************
* Go to Excel
**********************************************************************

    "Select cell top left
    CALL METHOD OF
        co_sheet
        'Cells'  = lo_cell
      EXPORTING
        #1       = 2
        #2       = 1.

    "Paste Clipboard
    " 1.) Select top left
    CALL METHOD OF
      lo_cell
      'SELECT'.

    " 2.) Paste
    CALL METHOD OF
      co_sheet
      'PASTE'.

    FREE OBJECT lo_cell.
    CALL FUNCTION 'CONTROL_FLUSH'
      EXCEPTIONS
        OTHERS = 3.

    lv_sheet_name = iv_sheet_name.
    REPLACE ALL OCCURRENCES OF '/' IN lv_sheet_name WITH '_'.
    SET PROPERTY OF co_sheet 'Name' = lv_sheet_name.

  ENDMETHOD.


  METHOD generate_output_table.
    DATA: ls_string TYPE string.
    DATA: ls_column TYPE abap_compdescr.
    DATA: lv_string TYPE string.

    FIELD-SYMBOLS: <fs_table_line>    TYPE any,
                   <fs_field_content> TYPE any.

**********************************************************************
* Create Paste Content
**********************************************************************
    LOOP AT it_table ASSIGNING <fs_table_line>.
      " fill the string with the contents of the fields
      LOOP AT io_ref_struc->components INTO ls_column.
        ASSIGN COMPONENT ls_column-name OF STRUCTURE <fs_table_line> TO <fs_field_content>.
        " Der erste Eintrag in einer Zeile sollte in Feld 1 geschrieben werden (Seperierung durch horizontal_tab.
        IF sy-tabix = 1.
          ls_string = <fs_field_content>.
          " > Erster Eintrag sollte mit den vorgegangenen getrennt durch horizontal_tab getrennt sein.
        ELSE.
          IF <fs_field_content> IS ASSIGNED.
            lv_string = <fs_field_content>.
            CONCATENATE ls_string lv_string INTO ls_string SEPARATED BY cl_abap_char_utilities=>horizontal_tab.
          ENDIF.
        ENDIF.
      ENDLOOP.
      APPEND ls_string TO et_output_table.
      CLEAR ls_string.
    ENDLOOP.
  ENDMETHOD.


  METHOD get_active_sheet.
    DATA: lo_oldsheet         TYPE ole2_object.
    DATA: lv_number_of_sheets TYPE i.

    " How many sheets are existing?
    GET PROPERTY OF co_sheets 'COUNT' = lv_number_of_sheets.

    " If the requested index is not existing, create a new one.
    IF iv_sheet_index <= lv_number_of_sheets.

      CALL METHOD OF
          co_excel
          'WORKSHEETS' = co_sheet
        EXPORTING
          #1           = iv_sheet_index.

      CALL METHOD OF
        co_sheet
        'Activate'.

      " if more are needed, generate new sheets
    ELSE.


      lo_oldsheet = co_sheet.

      FREE co_sheet.

      " This creates a sheet before the active one
      CALL METHOD OF
        co_sheets
          'Add' = co_sheet.

      " Move the previous sheet left of the current one
      " to keep the order
      CALL METHOD OF
        lo_oldsheet
        'MOVE'
        EXPORTING
          #1 = co_sheet.

      " Activate the sheet
      CALL METHOD OF
        co_sheet
        'Activate'.

    ENDIF.

  ENDMETHOD.


  METHOD get_excel.
    INCLUDE ole2incl.

    DATA: lo_workbooks              TYPE ole2_object.
    DATA: lo_workbook               TYPE ole2_object.
    DATA: lo_sheets                 TYPE ole2_object.
    DATA: lo_sheet                  TYPE ole2_object.
    DATA: lv_orig_nr_sheet_in_workb TYPE i.

**********************************************************************
* Generate Excel Instance/Reference
**********************************************************************
    "Create Excel Object
    CREATE OBJECT eo_excel 'excel.application'.

    SET PROPERTY OF eo_excel 'SheetsInNewWorkbook' = 1.

    "Display Excel instance
    SET PROPERTY OF eo_excel 'VISIBLE' = 1.

    "Create Workbooks
    CALL METHOD OF
      eo_excel
        'WORKBOOKS' = lo_workbooks.
    "Create Workboook
    CALL METHOD OF
      lo_workbooks
        'ADD' = lo_workbook.

    "Determine Sheets
    GET PROPERTY OF lo_workbook 'WORKSHEETS' = eo_sheets.
    SET PROPERTY OF eo_excel 'SheetsInNewWorkbook' = 3.
  ENDMETHOD.


  METHOD get_number_of_columns.
    DATA(l_counter) = 0.
    LOOP AT io_ref_struc->components INTO DATA(ls_column).
      rv_number_of_columns = rv_number_of_columns + 1.
    ENDLOOP.
  ENDMETHOD.


  METHOD output_to_excel.

    TYPE-POOLS ole2.

    DATA: lo_excel             TYPE ole2_object.                        "Instance of Excel
    DATA: lo_workbooks         TYPE ole2_object.                        "Reference to workbooks of one instance
    DATA: lo_workbook          TYPE ole2_object.                        "Reference to a workbook
    DATA: lo_sheet             TYPE ole2_object.                        "Reference to a sheet
    DATA: lo_sheets            TYPE ole2_object.                        "Reference to sheets
    DATA: lt_string            TYPE STANDARD TABLE OF ty_excel_cell.    "Generic String Table for Output
    DATA: lr_struc             TYPE REF TO cl_abap_structdescr.         "Description of it_table
    DATA: lr_dref              TYPE REF TO data.
    DATA: lv_number_of_columns TYPE i.

    FIELD-SYMBOLS: <fs_field_content> TYPE any,
                   <fs_table_line>    TYPE any,
                   <fs_struc>         TYPE any.

**********************************************************************
    "Prozessierung der Input Tabelle

    "Beschreibung der Input Datentabelle
    CREATE DATA lr_dref LIKE LINE OF it_table.
    ASSIGN lr_dref->* TO <fs_struc>.
    lr_struc ?= cl_abap_structdescr=>describe_by_data( <fs_struc> ).

    IF iv_group_condition IS INITIAL.
**********************************************************************
* Get an excel object with one sheet
**********************************************************************

      get_excel( IMPORTING eo_sheets = lo_sheets
                           eo_excel  = lo_excel ).

      get_active_sheet( EXPORTING iv_sheet_index = 1
                        CHANGING  co_excel       = lo_excel
                                  co_sheets      = lo_sheets
                                  co_sheet       = lo_sheet ).

**********************************************************************
* Create Header
**********************************************************************

      set_header( EXPORTING io_ref_struc = lr_struc
                  CHANGING  co_sheet     = lo_sheet ).

**********************************************************************
* Create Content
**********************************************************************

      generate_output_table( EXPORTING io_ref_struc    = lr_struc
                                       it_table        = it_table
                             IMPORTING et_output_table = lt_string ).

**********************************************************************
* Fill Sheet
**********************************************************************

      fill_sheet( EXPORTING io_ref_struc    = lr_struc
                            it_output_table = lt_string
                            iv_sheet_name   = iv_sheet_name
                  CHANGING  co_sheet        = lo_sheet ).

**********************************************************************
* Auto format columns
**********************************************************************

      CLEAR lv_number_of_columns.
      lv_number_of_columns = get_number_of_columns( lr_struc ).
      DO lv_number_of_columns TIMES.
        autoformat_column_width( EXPORTING iv_column_index = sy-index
                                 CHANGING  co_sheet        = lo_sheet ).
      ENDDO.

      " If each DSO information should be written to a seperate sheet
    ELSE.

      DATA: lt_table              TYPE REF TO data,
            lt_table2             TYPE REF TO data,
            lt_table3             TYPE REF TO data,
            lv_where_cond         TYPE string,
            lv_comperator         TYPE string,
            lv_num_of_lines        TYPE i,
            lv_cur_line_nr         TYPE i,
            lv_cur_dso_nr_minus_1 TYPE i,
            lv_sheet_name         TYPE string,
            lo_oldsheet           TYPE ole2_object.

      FIELD-SYMBOLS: <ft_table>           TYPE ANY TABLE, " ultimately a table with one entry for each DSO
                     <ft_table2>          TYPE ANY TABLE, " Table with all data to be printed sorted by DSO
                     <ft_whole_table>       TYPE ANY TABLE, " Table that contains all entries of one DSO

                     <fs_dso_of_new_line> TYPE any.

      " First we have to identify how many DSOs are in the resul

      " generate a working copy of the result table.
      CREATE DATA lt_table LIKE it_table.
      ASSIGN lt_table->* TO <ft_table>.
      CREATE DATA lt_table2 LIKE it_table.
      ASSIGN lt_table2->* TO <ft_whole_table>.
      CREATE DATA lt_table3 LIKE it_table.
      ASSIGN lt_table3->* TO <ft_table2>.

      <ft_table>  = it_table.
      <ft_table2> = it_table.

      " sort it by Group Condition
      SORT <ft_table>  BY (iv_group_condition).
      SORT <ft_table2> BY (iv_group_condition).

      " delete double entries with respect of group condition
      DELETE ADJACENT DUPLICATES FROM  <ft_table> COMPARING (iv_group_condition).

      "Number of Lines
      DESCRIBE TABLE <ft_table> LINES lv_num_of_lines.

**********************************************************************
* Generate Excel Instance/Reference
**********************************************************************
**********************************************************************
* Get an excel object with one sheet
**********************************************************************
      get_excel( IMPORTING eo_sheets = lo_sheets
                           eo_excel  = lo_excel ).

      lv_cur_line_nr  = 0.
      LOOP AT <ft_table> ASSIGNING FIELD-SYMBOL(<fs_dso>).
        " Number of current lines
        lv_cur_line_nr  = lv_cur_line_nr  + 1.
        ASSIGN COMPONENT iv_group_condition OF STRUCTURE <fs_dso> TO FIELD-SYMBOL(<fs_cur_line>).
        lv_sheet_name = <fs_cur_line>.
        REFRESH <ft_whole_table>.

        return_subtable( EXPORTING it_all_data  = it_table
                                   it_key_field = iv_group_condition
                                   it_key_value = lv_sheet_name
                         IMPORTING et_subtable  = <ft_whole_table> ).

**********************************************************************
* Get Active Sheet
**********************************************************************
        get_active_sheet( EXPORTING iv_sheet_index = lv_cur_line_nr
                          CHANGING  co_excel       = lo_excel
                                    co_sheets      = lo_sheets
                                    co_sheet       = lo_sheet ).

**********************************************************************
* Create Header
**********************************************************************
        set_header( EXPORTING io_ref_struc = lr_struc
                    CHANGING  co_sheet     = lo_sheet ).

**********************************************************************
* Create Paste Content
**********************************************************************
        CLEAR lt_string.

        generate_output_table( EXPORTING io_ref_struc    = lr_struc
                                         it_table        = <ft_whole_table>
                               IMPORTING et_output_table = lt_string ).

**********************************************************************
* Fill Sheet
**********************************************************************

        fill_sheet( EXPORTING io_ref_struc    = lr_struc
                              it_output_table = lt_string
                              iv_sheet_name   = lv_sheet_name
                    CHANGING  co_sheet        = lo_sheet ).

**********************************************************************
* Auto format columns
**********************************************************************
        CLEAR lv_number_of_columns.
        lv_number_of_columns = get_number_of_columns( lr_struc ).
        DO lv_number_of_columns TIMES.
          autoformat_column_width( EXPORTING iv_column_index = sy-index
                                   CHANGING  co_sheet        = lo_sheet ).
        ENDDO.
      ENDLOOP.
    ENDIF.
  ENDMETHOD.


  METHOD return_subtable.

    DATA: lv_comperator TYPE string.
    DATA: lv_where_cond TYPE string.

    FIELD-SYMBOLS: <fs_table_line> TYPE any.

    CLEAR lv_comperator.
    CONCATENATE '''' it_key_value '''' INTO lv_comperator.
    CONCATENATE it_key_field '=' lv_comperator INTO lv_where_cond SEPARATED BY space.

    "Split result table by DSO
    LOOP AT it_all_data ASSIGNING <fs_table_line> WHERE (lv_where_cond).
      INSERT <fs_table_line> INTO TABLE et_subtable.
    ENDLOOP.

  ENDMETHOD.


  METHOD set_header.
    DATA: lo_cell     TYPE ole2_object.
    DATA: lo_font     TYPE ole2_object.
    DATA: lo_interior TYPE ole2_object.

    DATA(l_counter) = 0.

    LOOP AT io_ref_struc->components INTO DATA(ls_column).
      l_counter = l_counter + 1.
      " Wählen der Entsprechenden Zelle
      CALL METHOD OF
          co_sheet
          'Cells'  = lo_cell
        EXPORTING
          #1       = 1
          #2       = l_counter.

      " Setzen des Schrifttyps
      SET PROPERTY OF lo_cell 'VALUE' = ls_column-name.
      CALL METHOD OF
        lo_cell
          'FONT' = lo_font.
      SET PROPERTY OF lo_font 'BOLD' = 1.

      " Setzen des Zelleninhalts
      CALL METHOD OF
        lo_cell
          'Interior' = lo_interior.
      SET PROPERTY OF lo_interior 'Color' = 15773696.

      FREE lo_interior.
      FREE lo_font.
      FREE lo_cell.

    ENDLOOP.

  ENDMETHOD.
ENDCLASS.
