#Requires AutoHotkey >=v2
#SingleInstance Force

^Esc::ExitApp

class Constants
{

  class ImageFile 
  {
    static _base_path := "C:\Users\vnsta\AppData\Roaming\ImageRecognition\"
    static _image_suffix := ".png"

    static image_select_all := this._create_path("ImageOfSelectAll")
    static image_column_seperater := this._create_path("ImageOfColumnSeperater")
    static image_navigation_panel := this._create_path("ImageOfNavigationPanel")
    static image_bolt_button := this._create_path("ImageOfBoltButton")
    
    static _create_path(image_name)
    {
      return this._base_path . image_name . this._image_suffix
    }
  }

  class Window 
  {
    static pricelist_name := "Vendor Name Pricelist.pdf - Adobe Acrobat Reader (64-bit)"
    static pricelist_class := "ahk_class AcrobatSDIWindow"
    static excel_no_table_name := "Excel"
    static excel_table_name := "331645 Lieferantenname.xlsx - Excel"
    static excel_class := "ahk_class XLMAIN"

    static pricelist_properties := Constants.Window.pricelist_name . " " . Constants.Window.pricelist_class
    static excel_no_table_properties := Constants.Window.excel_no_table_name . " " . Constants.Window.excel_class
    static excel_table_properties := Constants.Window.excel_table_name . " " . Constants.Window.excel_class

  }
}

Class Table 
{
  __New() 
  {
    this.max_row := 0
    this.max_column := 0
    this.table_columns := Array()
    
    this._column_separator := ";"
    this._row_separator := "`n"
  }

  _compute_csv_dimension(file_name)
  {
    reached_end_of_line := false
    loop parse FileRead(file_name), this._column_separator
    {
      if (not reached_end_of_line)
        this.max_column += 1
      
      if InStr(A_LoopField, this._row_separator)
      {
        reached_end_of_line := true
        this.max_row += 1
      }
    }
  }

  import_csv(file_name)
  {
    this._compute_csv_dimension(file_name)
    loop this.max_column
      this.table_columns.Push(Array())
    
    column := 1
    loop parse FileRead(file_name), this._column_separator . this._row_separator, "`r"
    {
      this.table_columns[column].Push(A_LoopField)
      column += 1
      if column > this.max_column
        column := 1
    }

    return this.table_columns
  }

  export_csv(location)
  {
    if FileExist(location) != ""
      FileDelete(location)
    
    csv_table := ""
    loop this.max_row
    {
      current_row := A_Index
      for column_array in this.table_columns
        csv_table .= column_array[current_row] . this._column_separator

      csv_table := SubStr(csv_table, 1, StrLen(csv_table) - 1) . this._row_separator
    }

    FileAppend(csv_table, location)
  }

  search_and_replace_column(search_term, replace_term, column, ignore_header := false)
  {
    column_array := this.table_columns[column]

    for cell_entry in column_array
    {
      if (ignore_header)
      {
        ignore_header := false
        continue
      }
        
      replacement := StrReplace(cell_entry, search_term, replace_term)
      column_array[A_Index] := replacement
    }
  }

  add_number_to_column_elements(number_to_add, column, ignore_header := false)
  {
    column_array := this.table_columns[column]

    for value in column_array
    {
      if (ignore_header)
      {
        ignore_header := false
        continue
      }

      if IsNumber(column_array[A_Index])
        column_array[A_Index] := Number(column_array[A_Index]) + number_to_add
    }
  }
}

try_open_pricelist()
{
  if (not WinActive(Constants.Window.pricelist_properties))
    Run("D:\Lieferanten\331645 Lieferantenname\Preisliste\Vendor Name Pricelist.pdf")
  WinWaitActive(Constants.Window.pricelist_name)
}

copy_pricelist()
{
  WinActivate(Constants.Window.pricelist_properties)
  WinWaitActive(Constants.Window.pricelist_name)
  A_Clipboard := ""
  Sleep 1000 ; wait for newline seperator to load (?)
  SendInput("{Ctrl Down}{a}{Ctrl Up}")
  SendInput("{Ctrl Down}{c}{Ctrl Up}")
  ClipWait(1)
  ; waits until every data entry is in clipboard
  If ClipWait(1) = false
  {
    MsgBox("No text appeared on clipboard after Ctrl+C and waiting for one second","Error",4144)
    Return
  }
  Loop 
  {
    OldClip := A_Clipboard
    Sleep 1000
  } Until A_Clipboard = OldClip
}

create_csv_file_from_clipboard(max_column, file_destination)
{
  csv_file := "" ;"sep=;`n"
  Loop Parse A_Clipboard, "`n", "`r"
  {
    char := ";"
    If (Mod(A_Index, max_column) = 0)
      char := "`n"
    csv_file .= A_LoopField . char
  }
  FileAppend(csv_file, file_destination)
}

create_xlsx_file_from_csv(csv_file_location_to_convert)
{
  xl := ComObject("Excel.Application") 
  xl.Workbooks.OpenText(csv_file_location_to_convert,,, 1)  ; 1: xlDelimited
  SplitPath(csv_file_location_to_convert, , &OutDir, , &OutNameNoExt)
  xl.ActiveWorkbook.SaveAs(OutDir . "\" . OutNameNoExt . ".xlsx", 51) ; https://learn.microsoft.com/de-de/office/vba/api/excel.xlfileformat
  xl.ActiveWorkbook.Close()
  xl.quit()
}

; try_open_xlsx_pricelist()
; {
;   if (not WinActive(Constants.Window.excel_table_properties))
;     Run("D:\Lieferanten\331645 Lieferantenname\Preisliste\331645 Lieferantenname.xlsx")
;   WinWaitActive(Constants.Window.excel_table_name)
; }

; format_xlsx_pricelist(max_column)
; {
;   find_and_click_image(Constants.ImageFile.image_select_all)
;   find_and_click_image(Constants.ImageFile.image_column_seperater, 2)
;   find_and_click_image(Constants.ImageFile.image_navigation_panel)
;   SendText("Z1S1:Z1S" . max_column)
;   SendInput("{Enter}")
;   find_and_click_image(Constants.ImageFile.image_bolt_button)
; }

format_xlsx_pricelist(max_column, xlsx_file_location)
{
  xl := ComObject("Excel.Application")
  xl.Workbooks.Open(xlsx_file_location)
  xl := ComObjActive("Excel.Application")
  cell_1 := xl.Sheets(1).Cells(1, 1)
  cell_2 := xl.Sheets(1).Cells(1, max_column)
  
  xl.Sheets(1).Range(cell_1, cell_2).Font.Bold := True
  xl.ActiveWorkbook.ActiveSheet.Cells.Select
  xl.Sheets(1).Cells.EntireColumn.AutoFit
  xl.Sheets(1).Cells(1, 1).Select

  xl.ActiveWorkbook.Save
  xl.ActiveWorkbook.Close
}

manipulate_csv(import_file, export_file)
{
  tb1 := Table()
  tb1.import_csv(import_file)
  tb1.search_and_replace_column(" " . "days", "", 7)
  tb1.add_number_to_column_elements(3, 7)
  tb1.search_and_replace_column(",", "", 6)
  tb1.search_and_replace_column(".", ",", 6)
  tb1.search_and_replace_column("vendor name", "331645", 1, true)
  tb1.export_csv(export_file)
}

main()
{
  max_column := 7
  file_destination := "D:\Lieferanten\331645 Lieferantenname\Preisliste\331645 Lieferantenname.csv"
  file_location := "D:\Lieferanten\331645 Lieferantenname\Preisliste\331645 Lieferantenname.csv"
  csv_file_location_to_convert := "D:\Lieferanten\331645 Lieferantenname\Preisliste\331645 Lieferantenname.csv"
  xlsx_file_location := "D:\Lieferanten\331645 Lieferantenname\Preisliste\331645 Lieferantenname.xlsx"

  SetTitleMatchMode(2)
  CoordMode("Pixel", "Window")
  CoordMode("Mouse", "Window")

  try_open_pricelist()
  copy_pricelist()
  WinClose(Constants.Window.pricelist_properties)
  create_csv_file_from_clipboard(max_column, file_destination) 
  create_xlsx_file_from_csv(csv_file_location_to_convert)
  format_xlsx_pricelist(max_column, xlsx_file_location)
  manipulate_csv(file_location, file_destination)
  MsgBox("Done", "Pricelist Conversion")
}

; auto-execute section
main()
ExitApp(0)