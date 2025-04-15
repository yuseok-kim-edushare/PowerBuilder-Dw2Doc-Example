$PBExportHeader$w_stringdialog.srw
forward
global type w_stringdialog from window
end type
type st_prompt from statictext within w_stringdialog
end type
type sle_name from singlelineedit within w_stringdialog
end type
type cb_ok from commandbutton within w_stringdialog
end type
type cb_cancel from commandbutton within w_stringdialog
end type
end forward

global type w_stringdialog from window
integer width = 1307
integer height = 556
boolean titlebar = true
string title = "New profile"
boolean controlmenu = true
windowtype windowtype = response!
long backcolor = 67108864
string icon = "AppIcon!"
boolean center = true
st_prompt st_prompt
sle_name sle_name
cb_ok cb_ok
cb_cancel cb_cancel
end type
global w_stringdialog w_stringdialog

forward prototypes
public function string wf_GetInput ()
public subroutine wf_ispassword (boolean as_ispassword)
end prototypes

public function string wf_GetInput ();return sle_name.Text;
end function

public subroutine wf_ispassword (boolean as_ispassword);sle_name.Password = as_ispassword
end subroutine

on w_stringdialog.create
this.st_prompt=create st_prompt
this.sle_name=create sle_name
this.cb_ok=create cb_ok
this.cb_cancel=create cb_cancel
this.Control[]={this.st_prompt,&
this.sle_name,&
this.cb_ok,&
this.cb_cancel}
end on

on w_stringdialog.destroy
destroy(this.st_prompt)
destroy(this.sle_name)
destroy(this.cb_ok)
destroy(this.cb_cancel)
end on

event open;str_dialogparams str_params
str_params = Message.PowerObjectParm

if IsValid(str_params) and not IsNull(str_params) then
	if str_params.s_title <> '' then this.Title = str_params.s_title
	if str_params.s_prompt <> '' then st_prompt.Text = str_params.s_prompt
	if str_params.s_content <> '' then sle_name.Text = str_params.s_content
	wf_ispassword(str_params.b_ispassword)
end if


end event

event resize;sle_name.width = newwidth - (st_prompt.x + st_prompt.width + 128)
sle_name.x = st_prompt.x + st_prompt.width + 16
end event

type st_prompt from statictext within w_stringdialog
integer x = 50
integer y = 140
integer width = 283
integer height = 68
integer textsize = -10
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
long textcolor = 33554432
long backcolor = 67108864
string text = "Name:"
boolean focusrectangle = false
end type

type sle_name from singlelineedit within w_stringdialog
integer x = 320
integer y = 124
integer width = 827
integer height = 112
integer taborder = 10
integer textsize = -10
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
long textcolor = 33554432
borderstyle borderstyle = stylelowered!
end type

type cb_ok from commandbutton within w_stringdialog
integer x = 709
integer y = 328
integer width = 402
integer height = 112
integer taborder = 20
integer textsize = -10
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
string text = "Ok"
boolean default = true
end type

event clicked;Message.longparm = 1

if sle_name.text = "" then
	sle_name.SetFocus()
end if

CloseWithReturn(Parent, sle_name.text)
end event

type cb_cancel from commandbutton within w_stringdialog
integer x = 178
integer y = 328
integer width = 402
integer height = 112
integer taborder = 10
integer textsize = -10
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
string text = "Cancel"
end type

event clicked;Message.longparm = 0
close(Parent)
end event

