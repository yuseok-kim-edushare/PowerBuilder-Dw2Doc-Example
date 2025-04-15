$PBExportHeader$w_dw2excelapp_settings.srw
forward
global type w_dw2excelapp_settings from window
end type
type st_s_openloggingdir from u_unthemedstatictext within w_dw2excelapp_settings
end type
type st_s_loggingdesc from statictext within w_dw2excelapp_settings
end type
type cbx_logging from checkbox within w_dw2excelapp_settings
end type
type st_s_logging from statictext within w_dw2excelapp_settings
end type
type st_s_properties from statictext within w_dw2excelapp_settings
end type
type cbx_color from checkbox within w_dw2excelapp_settings
end type
type st_s_bands from statictext within w_dw2excelapp_settings
end type
type lb_bands from listbox within w_dw2excelapp_settings
end type
type st_s_misaligneddescription from statictext within w_dw2excelapp_settings
end type
type st_s_misaligned from statictext within w_dw2excelapp_settings
end type
type cbx_misaligned from checkbox within w_dw2excelapp_settings
end type
type cb_cancel from commandbutton within w_dw2excelapp_settings
end type
type cb_ok from commandbutton within w_dw2excelapp_settings
end type
type cbx_buttons from checkbox within w_dw2excelapp_settings
end type
type cbx_bitmaps from checkbox within w_dw2excelapp_settings
end type
type cbx_ovals from checkbox within w_dw2excelapp_settings
end type
type cbx_rectangles from checkbox within w_dw2excelapp_settings
end type
type cbx_line from checkbox within w_dw2excelapp_settings
end type
type st_s_targets from statictext within w_dw2excelapp_settings
end type
type st_s_ythreshold from statictext within w_dw2excelapp_settings
end type
type em_ythreshold from editmask within w_dw2excelapp_settings
end type
type em_xthreshold from editmask within w_dw2excelapp_settings
end type
type st_s_thresholddescription from statictext within w_dw2excelapp_settings
end type
type st_s_thresholds from statictext within w_dw2excelapp_settings
end type
type st_s_xthreshold from statictext within w_dw2excelapp_settings
end type
type st_s_bandsdescription from statictext within w_dw2excelapp_settings
end type
end forward

global type w_dw2excelapp_settings from window
integer width = 2190
integer height = 1544
boolean titlebar = true
string title = "Conversion settings"
boolean controlmenu = true
windowtype windowtype = response!
long backcolor = 67108864
string icon = "Function!"
boolean center = true
st_s_openloggingdir st_s_openloggingdir
st_s_loggingdesc st_s_loggingdesc
cbx_logging cbx_logging
st_s_logging st_s_logging
st_s_properties st_s_properties
cbx_color cbx_color
st_s_bands st_s_bands
lb_bands lb_bands
st_s_misaligneddescription st_s_misaligneddescription
st_s_misaligned st_s_misaligned
cbx_misaligned cbx_misaligned
cb_cancel cb_cancel
cb_ok cb_ok
cbx_buttons cbx_buttons
cbx_bitmaps cbx_bitmaps
cbx_ovals cbx_ovals
cbx_rectangles cbx_rectangles
cbx_line cbx_line
st_s_targets st_s_targets
st_s_ythreshold st_s_ythreshold
em_ythreshold em_ythreshold
em_xthreshold em_xthreshold
st_s_thresholddescription st_s_thresholddescription
st_s_thresholds st_s_thresholds
st_s_xthreshold st_s_xthreshold
st_s_bandsdescription st_s_bandsdescription
end type
global w_dw2excelapp_settings w_dw2excelapp_settings

type variables
str_renderingsettings istr_settings

end variables

forward prototypes
public subroutine of_loaddata (str_renderingsettings astr_settings)
end prototypes

public subroutine of_loaddata (str_renderingsettings astr_settings);int i

cbx_bitmaps.checked = astr_settings.b_renderbitmap
cbx_line.checked = astr_settings.b_renderline
cbx_rectangles.checked = astr_settings.b_renderrectangle
cbx_ovals.checked = astr_settings.b_renderoval
cbx_buttons.checked = astr_settings.b_renderbutton
cbx_misaligned.Checked = astr_settings.b_ignoremisaligned
cbx_color.Checked = astr_settings.b_exportcolor
cbx_logging.Checked = astr_settings.b_logging

for i = 1 to UpperBound(astr_settings.s_bands)
	lb_bands.SetState(lb_bands.FindItem(astr_settings.s_bands[i], 0), true)
next

em_xthreshold.Text = string(astr_settings.i_xthreshold)
em_ythreshold.Text = string(astr_settings.i_ythreshold)
end subroutine

on w_dw2excelapp_settings.create
this.st_s_openloggingdir=create st_s_openloggingdir
this.st_s_loggingdesc=create st_s_loggingdesc
this.cbx_logging=create cbx_logging
this.st_s_logging=create st_s_logging
this.st_s_properties=create st_s_properties
this.cbx_color=create cbx_color
this.st_s_bands=create st_s_bands
this.lb_bands=create lb_bands
this.st_s_misaligneddescription=create st_s_misaligneddescription
this.st_s_misaligned=create st_s_misaligned
this.cbx_misaligned=create cbx_misaligned
this.cb_cancel=create cb_cancel
this.cb_ok=create cb_ok
this.cbx_buttons=create cbx_buttons
this.cbx_bitmaps=create cbx_bitmaps
this.cbx_ovals=create cbx_ovals
this.cbx_rectangles=create cbx_rectangles
this.cbx_line=create cbx_line
this.st_s_targets=create st_s_targets
this.st_s_ythreshold=create st_s_ythreshold
this.em_ythreshold=create em_ythreshold
this.em_xthreshold=create em_xthreshold
this.st_s_thresholddescription=create st_s_thresholddescription
this.st_s_thresholds=create st_s_thresholds
this.st_s_xthreshold=create st_s_xthreshold
this.st_s_bandsdescription=create st_s_bandsdescription
this.Control[]={this.st_s_openloggingdir,&
this.st_s_loggingdesc,&
this.cbx_logging,&
this.st_s_logging,&
this.st_s_properties,&
this.cbx_color,&
this.st_s_bands,&
this.lb_bands,&
this.st_s_misaligneddescription,&
this.st_s_misaligned,&
this.cbx_misaligned,&
this.cb_cancel,&
this.cb_ok,&
this.cbx_buttons,&
this.cbx_bitmaps,&
this.cbx_ovals,&
this.cbx_rectangles,&
this.cbx_line,&
this.st_s_targets,&
this.st_s_ythreshold,&
this.em_ythreshold,&
this.em_xthreshold,&
this.st_s_thresholddescription,&
this.st_s_thresholds,&
this.st_s_xthreshold,&
this.st_s_bandsdescription}
end on

on w_dw2excelapp_settings.destroy
destroy(this.st_s_openloggingdir)
destroy(this.st_s_loggingdesc)
destroy(this.cbx_logging)
destroy(this.st_s_logging)
destroy(this.st_s_properties)
destroy(this.cbx_color)
destroy(this.st_s_bands)
destroy(this.lb_bands)
destroy(this.st_s_misaligneddescription)
destroy(this.st_s_misaligned)
destroy(this.cbx_misaligned)
destroy(this.cb_cancel)
destroy(this.cb_ok)
destroy(this.cbx_buttons)
destroy(this.cbx_bitmaps)
destroy(this.cbx_ovals)
destroy(this.cbx_rectangles)
destroy(this.cbx_line)
destroy(this.st_s_targets)
destroy(this.st_s_ythreshold)
destroy(this.em_ythreshold)
destroy(this.em_xthreshold)
destroy(this.st_s_thresholddescription)
destroy(this.st_s_thresholds)
destroy(this.st_s_xthreshold)
destroy(this.st_s_bandsdescription)
end on

event open;str_renderingsettings lstr_settings

lstr_settings = Message.PowerObjectparm

istr_settings = lstr_settings
of_loaddata(lstr_settings)
end event

type st_s_openloggingdir from u_unthemedstatictext within w_dw2excelapp_settings
integer x = 1829
integer y = 880
integer textsize = -8
string pointer = "Hyperlink!"
long textcolor = 134217856
long backcolor = 553648127
string text = "Open."
end type

event clicked;call super::clicked;Run("explorer.exe .\logs\")
end event

type st_s_loggingdesc from statictext within w_dw2excelapp_settings
integer x = 1371
integer y = 824
integer width = 745
integer height = 112
integer textsize = -8
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
long textcolor = 33554432
long backcolor = 67108864
string text = "Log files are saved on the program~'s directory."
boolean focusrectangle = false
end type

type cbx_logging from checkbox within w_dw2excelapp_settings
integer x = 1371
integer y = 948
integer width = 448
integer height = 92
integer textsize = -9
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
long textcolor = 33554432
long backcolor = 67108864
string text = "Enable Logging"
end type

event clicked;istr_settings.b_logging = Checked
end event

type st_s_logging from statictext within w_dw2excelapp_settings
integer x = 1371
integer y = 736
integer width = 571
integer height = 76
integer textsize = -11
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
long textcolor = 33554432
long backcolor = 67108864
string text = "Logging"
boolean focusrectangle = false
end type

type st_s_properties from statictext within w_dw2excelapp_settings
integer x = 1371
integer y = 528
integer width = 571
integer height = 76
integer textsize = -11
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
long textcolor = 33554432
long backcolor = 67108864
string text = "Property targets"
boolean focusrectangle = false
end type

type cbx_color from checkbox within w_dw2excelapp_settings
integer x = 1371
integer y = 608
integer width = 338
integer height = 92
integer textsize = -9
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
long textcolor = 33554432
long backcolor = 67108864
string text = "Color"
end type

event clicked;istr_settings.b_exportcolor = Checked
end event

type st_s_bands from statictext within w_dw2excelapp_settings
integer x = 64
integer y = 860
integer width = 571
integer height = 76
integer textsize = -11
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
long textcolor = 33554432
long backcolor = 67108864
string text = "Bands"
boolean focusrectangle = false
end type

type lb_bands from listbox within w_dw2excelapp_settings
integer x = 64
integer y = 1084
integer width = 722
integer height = 248
integer taborder = 20
integer textsize = -10
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
long textcolor = 33554432
boolean sorted = false
boolean multiselect = true
string item[] = {"header","summary","footer"}
borderstyle borderstyle = stylelowered!
end type

type st_s_misaligneddescription from statictext within w_dw2excelapp_settings
integer x = 64
integer y = 596
integer width = 1129
integer height = 120
integer textsize = -8
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
long textcolor = 33554432
long backcolor = 67108864
string text = "When objects don~'t align on Y and/or overlap other objects, they are converted into floating objects."
boolean focusrectangle = false
end type

type st_s_misaligned from statictext within w_dw2excelapp_settings
integer x = 64
integer y = 512
integer width = 549
integer height = 76
integer textsize = -11
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
long textcolor = 33554432
long backcolor = 67108864
string text = "Misaligned objects"
boolean focusrectangle = false
end type

type cbx_misaligned from checkbox within w_dw2excelapp_settings
integer x = 64
integer y = 732
integer width = 736
integer height = 92
integer textsize = -9
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
long textcolor = 33554432
long backcolor = 67108864
string text = "Skip misaligned objects"
end type

event clicked;istr_settings.b_ignoremisaligned = Checked
end event

type cb_cancel from commandbutton within w_dw2excelapp_settings
integer x = 1723
integer y = 1280
integer width = 402
integer height = 112
integer taborder = 20
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

type cb_ok from commandbutton within w_dw2excelapp_settings
integer x = 1285
integer y = 1280
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
end type

event clicked;int i
string ls_bands[]

for i = 1 to lb_bands.TotalItems()
	if lb_bands.State(i) = 1 then
		ls_bands[UpperBound(ls_bands) + 1] = lb_bands.Item[i]
	end if
next

istr_settings.s_bands = ls_bands

CloseWithReturn(parent, istr_settings)
end event

type cbx_buttons from checkbox within w_dw2excelapp_settings
integer x = 1760
integer y = 276
integer width = 338
integer height = 92
integer textsize = -9
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
long textcolor = 33554432
long backcolor = 67108864
string text = "Buttons"
end type

event clicked;istr_settings.b_renderbutton = Checked
end event

type cbx_bitmaps from checkbox within w_dw2excelapp_settings
integer x = 1760
integer y = 184
integer width = 325
integer height = 92
integer textsize = -9
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
long textcolor = 33554432
long backcolor = 67108864
string text = "Bitmaps"
end type

event clicked;istr_settings.b_renderbitmap = Checked
end event

type cbx_ovals from checkbox within w_dw2excelapp_settings
integer x = 1371
integer y = 372
integer width = 338
integer height = 92
integer textsize = -9
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
long textcolor = 33554432
long backcolor = 67108864
string text = "Ovals"
end type

event clicked;istr_settings.b_renderoval = Checked
end event

type cbx_rectangles from checkbox within w_dw2excelapp_settings
integer x = 1371
integer y = 276
integer width = 338
integer height = 92
integer textsize = -9
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
long textcolor = 33554432
long backcolor = 67108864
string text = "Rectangles"
end type

event clicked;istr_settings.b_renderrectangle = Checked
end event

type cbx_line from checkbox within w_dw2excelapp_settings
integer x = 1371
integer y = 184
integer width = 325
integer height = 92
integer textsize = -9
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
long textcolor = 33554432
long backcolor = 67108864
string text = "Lines"
end type

event clicked;istr_settings.b_renderline = Checked
end event

type st_s_targets from statictext within w_dw2excelapp_settings
integer x = 1371
integer y = 76
integer width = 571
integer height = 76
integer textsize = -11
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
long textcolor = 33554432
long backcolor = 67108864
string text = "Conversion targets"
boolean focusrectangle = false
end type

type st_s_ythreshold from statictext within w_dw2excelapp_settings
integer x = 434
integer y = 308
integer width = 402
integer height = 64
integer textsize = -8
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
long textcolor = 33554432
long backcolor = 67108864
string text = "Y Threshold"
boolean focusrectangle = false
end type

type em_ythreshold from editmask within w_dw2excelapp_settings
integer x = 434
integer y = 372
integer width = 343
integer height = 92
integer taborder = 10
integer textsize = -8
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
long textcolor = 33554432
string text = "none"
boolean displayonly = true
borderstyle borderstyle = stylelowered!
string mask = "#"
boolean spin = true
double increment = 1
end type

event modified;If not IsNumber(Text) then
	MessageBox("Error", "Invalid input")
end if

istr_settings.i_ythreshold = Integer(Text)
end event

type em_xthreshold from editmask within w_dw2excelapp_settings
integer x = 64
integer y = 372
integer width = 343
integer height = 92
integer taborder = 10
integer textsize = -8
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
long textcolor = 33554432
string text = "none"
boolean displayonly = true
borderstyle borderstyle = stylelowered!
string mask = "#"
boolean spin = true
double increment = 1
end type

event modified;If not IsNumber(Text) then
	MessageBox("Error", "Invalid input")
end if

istr_settings.i_xthreshold = Integer(Text)
end event

type st_s_thresholddescription from statictext within w_dw2excelapp_settings
integer x = 64
integer y = 156
integer width = 1129
integer height = 148
integer textsize = -8
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
long textcolor = 33554432
long backcolor = 67108864
string text = "Defines how far an object can be from the rest before it~'s placed on another row/column."
boolean focusrectangle = false
end type

type st_s_thresholds from statictext within w_dw2excelapp_settings
integer x = 64
integer y = 76
integer width = 402
integer height = 64
integer textsize = -11
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
long textcolor = 33554432
long backcolor = 67108864
string text = "Thresholds"
boolean focusrectangle = false
end type

type st_s_xthreshold from statictext within w_dw2excelapp_settings
integer x = 64
integer y = 308
integer width = 402
integer height = 64
integer textsize = -8
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
long textcolor = 33554432
long backcolor = 67108864
string text = "X Threshold"
boolean focusrectangle = false
end type

type st_s_bandsdescription from statictext within w_dw2excelapp_settings
integer x = 64
integer y = 948
integer width = 750
integer height = 152
integer textsize = -8
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
long textcolor = 33554432
long backcolor = 67108864
string text = "DataWindow bands you want to include when exporting."
boolean focusrectangle = false
end type

