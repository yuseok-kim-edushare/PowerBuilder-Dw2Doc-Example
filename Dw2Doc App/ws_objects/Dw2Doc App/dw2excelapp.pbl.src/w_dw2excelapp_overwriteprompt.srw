$PBExportHeader$w_dw2excelapp_overwriteprompt.srw
forward
global type w_dw2excelapp_overwriteprompt from window
end type
type cb_cancel from commandbutton within w_dw2excelapp_overwriteprompt
end type
type cb_ok from commandbutton within w_dw2excelapp_overwriteprompt
end type
type st_s_sheetname from statictext within w_dw2excelapp_overwriteprompt
end type
type sle_sheetname from singlelineedit within w_dw2excelapp_overwriteprompt
end type
type rb_append from radiobutton within w_dw2excelapp_overwriteprompt
end type
type rb_overwrite from radiobutton within w_dw2excelapp_overwriteprompt
end type
type st_s_intro from statictext within w_dw2excelapp_overwriteprompt
end type
type gb_1 from groupbox within w_dw2excelapp_overwriteprompt
end type
end forward

global type w_dw2excelapp_overwriteprompt from window
integer width = 1719
integer height = 772
boolean titlebar = true
string title = "File already exists"
boolean controlmenu = true
windowtype windowtype = response!
long backcolor = 67108864
string icon = "AppIcon!"
boolean center = true
cb_cancel cb_cancel
cb_ok cb_ok
st_s_sheetname st_s_sheetname
sle_sheetname sle_sheetname
rb_append rb_append
rb_overwrite rb_overwrite
st_s_intro st_s_intro
gb_1 gb_1
end type
global w_dw2excelapp_overwriteprompt w_dw2excelapp_overwriteprompt

on w_dw2excelapp_overwriteprompt.create
this.cb_cancel=create cb_cancel
this.cb_ok=create cb_ok
this.st_s_sheetname=create st_s_sheetname
this.sle_sheetname=create sle_sheetname
this.rb_append=create rb_append
this.rb_overwrite=create rb_overwrite
this.st_s_intro=create st_s_intro
this.gb_1=create gb_1
this.Control[]={this.cb_cancel,&
this.cb_ok,&
this.st_s_sheetname,&
this.sle_sheetname,&
this.rb_append,&
this.rb_overwrite,&
this.st_s_intro,&
this.gb_1}
end on

on w_dw2excelapp_overwriteprompt.destroy
destroy(this.cb_cancel)
destroy(this.cb_ok)
destroy(this.st_s_sheetname)
destroy(this.sle_sheetname)
destroy(this.rb_append)
destroy(this.rb_overwrite)
destroy(this.st_s_intro)
destroy(this.gb_1)
end on

event open;rb_overwrite.Checked = true
rb_overwrite.event Clicked()
end event

type cb_cancel from commandbutton within w_dw2excelapp_overwriteprompt
integer x = 1225
integer y = 552
integer width = 402
integer height = 112
integer taborder = 40
integer textsize = -10
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
string text = "Cancel"
end type

event clicked;Close(Parent)
end event

type cb_ok from commandbutton within w_dw2excelapp_overwriteprompt
integer x = 777
integer y = 552
integer width = 402
integer height = 112
integer taborder = 30
integer textsize = -10
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
string text = "Ok"
end type

event clicked;str_overwritepromptresult lstr_result

if rb_append.Checked then
	if Trim(sle_sheetname.Text) = "" then
		sle_sheetname.SetFocus( )
		return
	end if
	lstr_result.s_sheetname = sle_sheetname.Text
end if

lstr_result.b_overwrite = rb_overwrite.Checked

CloseWithReturn(parent, lstr_result)
end event

type st_s_sheetname from statictext within w_dw2excelapp_overwriteprompt
integer x = 160
integer y = 376
integer width = 352
integer height = 64
integer textsize = -10
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
long textcolor = 33554432
long backcolor = 67108864
string text = "Sheet name:"
boolean focusrectangle = false
end type

type sle_sheetname from singlelineedit within w_dw2excelapp_overwriteprompt
integer x = 503
integer y = 352
integer width = 617
integer height = 112
integer taborder = 20
integer textsize = -10
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
long textcolor = 33554432
borderstyle borderstyle = stylelowered!
end type

type rb_append from radiobutton within w_dw2excelapp_overwriteprompt
integer x = 73
integer y = 260
integer width = 567
integer height = 80
integer textsize = -10
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
long textcolor = 33554432
long backcolor = 67108864
string text = "Append as a sheet"
end type

event clicked;st_s_sheetname.Visible = Checked
sle_sheetname.Visible = Checked
end event

type rb_overwrite from radiobutton within w_dw2excelapp_overwriteprompt
integer x = 73
integer y = 164
integer width = 402
integer height = 80
integer textsize = -10
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
long textcolor = 33554432
long backcolor = 67108864
string text = "Overwrite"
end type

event clicked;st_s_sheetname.Visible = rb_append.Checked
sle_sheetname.Visible = rb_append.Checked
end event

type st_s_intro from statictext within w_dw2excelapp_overwriteprompt
integer x = 73
integer y = 60
integer width = 1294
integer height = 64
integer textsize = -10
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
long textcolor = 33554432
long backcolor = 67108864
string text = "The selected file already exists. How to proceed?"
boolean focusrectangle = false
end type

type gb_1 from groupbox within w_dw2excelapp_overwriteprompt
boolean visible = false
integer x = 9
integer y = 68
integer width = 1522
integer height = 516
integer taborder = 10
integer textsize = -10
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
long textcolor = 33554432
long backcolor = 67108864
string text = "none"
end type

