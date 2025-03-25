#Requires AutoHotkey >=v2
#SingleInstance Force
#Include DateParseV2.ahk

^Esc::ExitApp

class Constants
{

  class ImageFile 
  {
    static _base_path := "ImageRecognition\"
    static _image_suffix := ".png"

    static image_delivery_date_one := this._create_path("ImageOfColumn1")
    static image_delivery_date_two := this._create_path("ImageOfColumn2")
    static image_real_delivery_date := this._create_path("ImageOfColumn3")
    
    static _create_path(image_name)
    {
      return this._base_path . image_name . this._image_suffix
    }
  }

  class Window 
  {
    static excel_table_name := "Lieferdatum Vergleich.xlsx - Excel"
    static excel_class := "ahk_class XLMAIN"

    static excel_table_properties := Constants.Window.excel_table_name . " " . Constants.Window.excel_class
  }
}

find_and_click_image(image_file, mouse_click_number := 1, off_set_x := 0, off_set_y := 0, search_until_found := true)
{
  found_x := 0
  found_y := 0

  loop
  {
    image_found := ImageSearch(&found_x, &found_y, 0, 0, 1920, 1040, image_file)
    If (image_found = 1)
    {
      center_imgage_search_coords(image_file, &found_x, &found_y)
      Click(found_x + off_set_x, found_y + off_set_y, "Left", mouse_click_number)
      Sleep 50
      break
    }
  } until (not search_until_found)
}

;UTILITIES---------------------------------------------------------
center_imgage_search_coords(file, &coord_x, &coord_y)
{
  _gui := Gui()
  pic := _gui.Add("Picture",, file)

  width := 0
  height := 0
	pic.GetPos(,, &width, &height)
	coord_x += width // 2
	coord_y += height // 2
}
;-------------------------------------------------------------------

copy_and_compare_delivery_dates()
{
  mouse_clicks:= 2
  off_set_y := 20 * A_Index

  A_Clipboard := ""
  find_and_click_image(Constants.ImageFile.image_delivery_date_one, mouse_clicks,, off_set_y)
  SendInput("{Ctrl Down}{a}{Ctrl Up}")
  SendInput("{Ctrl Down}{c}{Ctrl Up}")
  ClipWait(0.5)
  delivery_date_one := A_Clipboard
  if SubStr(delivery_date_one, 3, 1) != "/" ; exit app if no cells with dates exists
  {
    SendInput("{Esc}")
    ExitApp
  }
  A_Clipboard := ""
  find_and_click_image(Constants.ImageFile.image_delivery_date_two, mouse_clicks,, off_set_y)
  SendInput("{Ctrl Down}{a}{Ctrl Up}")
  SendInput("{Ctrl Down}{c}{Ctrl Up}")
  ClipWait()
  delivery_date_two := A_Clipboard
  
  real_delivery_date := compare_delivery_dates(delivery_date_one, delivery_date_two)

  return real_delivery_date
}

compare_delivery_dates(date_one, date_two)
{
  date_one := DateParse(date_one)
  date_two := DateParse(date_two)
  if Number(date_one) > Number(date_two)
    result := date_one
  else
    result := date_two

    result := FormatTime(result, "dd/MM/yyyy")
  return result
}

insert_delivery_date(real_delivery_date)
{
  mouse_clicks:= 2
  off_set_y := 20 * A_Index

  find_and_click_image(Constants.ImageFile.image_real_delivery_date, mouse_clicks,, off_set_y)
  SendText(real_delivery_date)
  SendInput("{Enter}")
}


main()
{
  SetTitleMatchMode(2)
  CoordMode("Pixel", "Window")
  CoordMode("Mouse", "Window")
  loop 
  {
    real_delivery_date:= copy_and_compare_delivery_dates()
    insert_delivery_date(real_delivery_date)
  }
}

; auto-execute section
KeyWait("F3", "D")
main()
ExitApp(0)
