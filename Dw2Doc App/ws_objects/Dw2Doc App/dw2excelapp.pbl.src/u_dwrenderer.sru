$PBExportHeader$u_dwrenderer.sru
forward
global type u_dwrenderer from userobject
end type
type pb_opensettings from picturebutton within u_dwrenderer
end type
type ddlb_column from dropdownlistbox within u_dwrenderer
end type
type hpb_exportprogress from hprogressbar within u_dwrenderer
end type
type dw_host from datawindow within u_dwrenderer
end type
type st_t_dwstyle from statictext within u_dwrenderer
end type
type cb_export from commandbutton within u_dwrenderer
end type
type st_t_column from statictext within u_dwrenderer
end type
type st_t_value from statictext within u_dwrenderer
end type
type cbx_usefilter from checkbox within u_dwrenderer
end type
type ddlb_operator from dropdownlistbox within u_dwrenderer
end type
type st_t_operator from statictext within u_dwrenderer
end type
type pb_filter from picturebutton within u_dwrenderer
end type
type ddlb_dwstyle from dropdownlistbox within u_dwrenderer
end type
type sle_value from singlelineedit within u_dwrenderer
end type
end forward

global type u_dwrenderer from userobject
integer width = 3296
integer height = 1876
event ue_postresize ( )
event ue_filtercontrolsvisible ( boolean ab_visible )
event ue_progressbarvisible ( boolean ab_visible )
event ue_exportprogress ( )
pb_opensettings pb_opensettings
ddlb_column ddlb_column
hpb_exportprogress hpb_exportprogress
dw_host dw_host
st_t_dwstyle st_t_dwstyle
cb_export cb_export
st_t_column st_t_column
st_t_value st_t_value
cbx_usefilter cbx_usefilter
ddlb_operator ddlb_operator
st_t_operator st_t_operator
pb_filter pb_filter
ddlb_dwstyle ddlb_dwstyle
sle_value sle_value
end type
global u_dwrenderer u_dwrenderer

type variables
constant string IS_SETTINGSFILE = ".\config\render.ini"

str_renderingsettings istr_defaultsettings
str_renderingsettings istr_currentsettings

string is_datawindows[] = {"d_order_detail_list", "d_person_list", "d_product_detail"}

int ii_SelectedDwIndex
n_cst_dw2excelconverter ino_converter
end variables

forward prototypes
public subroutine of_updatefilterdefinition ()
public subroutine of_init ()
public subroutine of_exportxlsx (powerobject ao_progresscallback)
public subroutine of_resize (long al_newwidth, long al_newheight)
public function str_renderingsettings of_loadsettings ()
public subroutine of_savesettings (readonly str_renderingsettings astr_settings)
end prototypes

event ue_filtercontrolsvisible(boolean ab_visible);ddlb_column.Visible = ab_visible
sle_value.Visible = ab_visible
st_t_column.Visible = ab_visible
st_t_value.Visible = ab_visible
pb_filter.Visible = ab_visible
st_t_operator.Visible = ab_visible
ddlb_operator.Visible = ab_visible
end event

event ue_progressbarvisible(boolean ab_visible);//st_maxrecords.Visible = ab_visible
hpb_exportprogress.Visible = ab_visible

//st_maxrecords.SetPosition(ToTop!)
hpb_exportprogress.SetPosition(Behind!)
hpb_exportprogress.Position = 0
//st_maxrecords.backcolor = -2W
end event

event ue_exportprogress();long ll_current
long ll_total

ll_current = Message.WordParm
ll_total = Message.LongParm

//st_MaxRecords.Text = string(ll_current) + "/" + string(ll_total)

hpb_exportprogress.Position = double(ll_current) / double(ll_total) * 100.0
yield()
end event

public subroutine of_updatefilterdefinition ();ddlb_column.Reset()
sle_value.Text = ""

string ls_columns []
Long ll_Loop
String ls_ColName

FOR ll_Loop = 1 TO long(dw_host.object.datawindow.column.count)
	ls_ColName = dw_host.Describe("#" + String( ll_Loop ) + ".Name")
	IF Long( dw_host.Describe( ls_ColName + ".Visible") ) > 0 THEN
		ls_columns[ UpperBound(ls_columns[]) + 1] = ls_ColName
	END IF
next

int i
for i = 1 to UpperBound(ls_columns) 
	ddlb_column.AddItem(ls_columns[i])
next

end subroutine

public subroutine of_init ();string ls_initerror

SetNull(ls_initerror)

ino_converter = create n_cst_dw2excelconverter
ino_converter.of_init(ls_initerror)

if not IsNull(ls_initerror) then
	MessageBox("Error", "Error initializing converter: " + ls_initerror)
	Close(Parent)
	return
end if

istr_currentsettings = of_loadsettings( )

istr_defaultsettings.i_xthreshold = 5
istr_defaultsettings.i_ythreshold = 5
istr_defaultsettings.b_renderline = true
istr_defaultsettings.b_renderrectangle = true
istr_defaultsettings.b_renderoval = true
istr_defaultsettings.b_renderbitmap = true
istr_defaultsettings.b_renderbutton = true

ddlb_dwstyle.SelectItem(1)
ddlb_dwstyle.event SelectionChanged(1)

event ue_progressbarvisible(false)
event ue_filterControlsVisible(false)



cbx_usefilter.checked = false
cbx_usefilter.event clicked()

end subroutine

public subroutine of_exportxlsx (powerobject ao_progresscallback);string ls_sheetname
string ls_error
string ls_path
string ls_filename
int li_ret
boolean lb_append

str_overwritepromptresult lstr_result
SetNull(ls_error)
cb_export.Enabled = false

ls_filename = is_DataWindows[ii_SelectedDwIndex]
if ls_filename = "" then
	ls_filename = "out"
end if

ls_path = GS_APPDIRECTORY + "\" + ls_filename

li_ret = GetFileSaveName("File save location", ls_path, ls_filename,  "txt",  "Excel Workbook (*.xslx),*.xlsx", "", 9283)

// using a WHILE instead of IF because we want to break out if the dialog is Canceled
do while (li_ret = 1)
	SetPointer(Hourglass!)
	
	if FileExists(ls_path) then
		Open(w_dw2excelapp_overwriteprompt)
		lstr_result = Message.Powerobjectparm
		if not IsValid(lstr_result) then
			exit
		end if
		lb_append = not lstr_result.b_overwrite
	end if
	
	if IsNull(ls_error) then
		event ue_progressbarvisible(true)
		ino_converter.istr_settings = istr_currentsettings
		ino_converter.of_convert(&
				dw_host,&
				ls_path,&
				lb_append,&
				lstr_result.s_sheetname,&
				ls_error,&
				ao_progressCallback,&
				"ue_ExportProgress")
		event ue_progressbarvisible(false)
	end if
	if not IsNull(ls_error) then 
		MessageBox("Could not export", ls_error)
	else
		MessageBox("Success", "Export complete")
	end if
	exit
loop

cb_export.Enabled = true
end subroutine

public subroutine of_resize (long al_newwidth, long al_newheight);dw_host.width = al_newwidth
dw_host.height = al_newheight - dw_host.y - 60
hpb_exportprogress.width = al_newwidth
hpb_exportprogress.height = 60
hpb_exportprogress.y = dw_host.y + dw_host.height - hpb_exportprogress.height


cb_export.x = al_newwidth - cb_export.width
pb_opensettings.x = cb_export.x - pb_opensettings.width - 16
end subroutine

public function str_renderingsettings of_loadsettings ();str_renderingsettings lstr_settings
string ls_bands

lstr_settings = istr_defaultsettings

lstr_settings.i_xthreshold = Integer(ProfileString(IS_SETTINGSFILE, "Threshold", "x_threshold", "5"))
lstr_settings.i_ythreshold = Integer(ProfileString(IS_SETTINGSFILE, "Threshold", "y_threshold", "5"))
lstr_settings.b_renderbitmap = 		 ProfileString(IS_SETTINGSFILE, "Targets", "bitmap", "true") = "true"
lstr_settings.b_renderbutton = 		 ProfileString(IS_SETTINGSFILE, "Targets", "button", "true") = "true"
lstr_settings.b_renderline = 			 ProfileString(IS_SETTINGSFILE, "Targets", "line", "true") = "true"
lstr_settings.b_renderrectangle = 	 ProfileString(IS_SETTINGSFILE, "Targets", "rectangle", "true") = "true"
lstr_settings.b_renderoval = 			 ProfileString(IS_SETTINGSFILE, "Targets", "oval", "true") = "true"
lstr_settings.b_exportcolor =			 ProfileString(IS_SETTINGSFILE, "Properties", "color", "true") = "true"
lstr_settings.b_logging = 				 ProfileString(IS_SETTINGSFILE, "Logging", "enabled", "true") = "true"

ls_bands = ProfileString(IS_SETTINGSFILE, "Targets", "bands", "")

ino_converter.ino_StringExtensions.Split(ls_bands, ",", ref lstr_settings.s_bands)

return lstr_settings
end function

public subroutine of_savesettings (readonly str_renderingsettings astr_settings);int file
string ls_bands

file = FileOpen(IS_SETTINGSFILE, TextMode!, Write!, LockReadWrite!)

if file = -1 then
	MessageBox("Error", "Could not open file for writing")
	return
end if

FileClose(file)

ls_bands = ino_converter.ino_stringextensions.Join(astr_settings.s_bands, ",")

SetProfileString(IS_SETTINGSFILE, "Threshold", "x_threshold", string(astr_settings.i_xthreshold))
SetProfileString(IS_SETTINGSFILE, "Threshold", "y_threshold", string(astr_settings.i_ythreshold))
SetProfileString(IS_SETTINGSFILE, "Targets", "line", string(astr_settings.b_renderline))
SetProfileString(IS_SETTINGSFILE, "Targets", "rectangle", string(astr_settings.b_renderrectangle))
SetProfileString(IS_SETTINGSFILE, "Targets", "oval", string(astr_settings.b_renderoval))
SetProfileString(IS_SETTINGSFILE, "Targets", "button", string(astr_settings.b_renderbutton))
SetProfileString(IS_SETTINGSFILE, "Targets", "bitmap", string(astr_settings.b_renderbitmap))
SetProfileString(IS_SETTINGSFILE, "Targets", "bands", ls_bands)
SetProfileString(IS_SETTINGSFILE, "Properties", "color", string(astr_settings.b_exportcolor))
SetProfileString(IS_SETTINGSFILE, "Logging", "enabled", string(astr_settings.b_logging))
end subroutine

on u_dwrenderer.create
this.pb_opensettings=create pb_opensettings
this.ddlb_column=create ddlb_column
this.hpb_exportprogress=create hpb_exportprogress
this.dw_host=create dw_host
this.st_t_dwstyle=create st_t_dwstyle
this.cb_export=create cb_export
this.st_t_column=create st_t_column
this.st_t_value=create st_t_value
this.cbx_usefilter=create cbx_usefilter
this.ddlb_operator=create ddlb_operator
this.st_t_operator=create st_t_operator
this.pb_filter=create pb_filter
this.ddlb_dwstyle=create ddlb_dwstyle
this.sle_value=create sle_value
this.Control[]={this.pb_opensettings,&
this.ddlb_column,&
this.hpb_exportprogress,&
this.dw_host,&
this.st_t_dwstyle,&
this.cb_export,&
this.st_t_column,&
this.st_t_value,&
this.cbx_usefilter,&
this.ddlb_operator,&
this.st_t_operator,&
this.pb_filter,&
this.ddlb_dwstyle,&
this.sle_value}
end on

on u_dwrenderer.destroy
destroy(this.pb_opensettings)
destroy(this.ddlb_column)
destroy(this.hpb_exportprogress)
destroy(this.dw_host)
destroy(this.st_t_dwstyle)
destroy(this.cb_export)
destroy(this.st_t_column)
destroy(this.st_t_value)
destroy(this.cbx_usefilter)
destroy(this.ddlb_operator)
destroy(this.st_t_operator)
destroy(this.pb_filter)
destroy(this.ddlb_dwstyle)
destroy(this.sle_value)
end on

event constructor;of_init()



end event

event destructor;destroy ino_converter
end event

type pb_opensettings from picturebutton within u_dwrenderer
integer x = 2747
integer y = 92
integer width = 110
integer height = 96
integer taborder = 20
integer textsize = -10
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
string picturename = "Preferences1!"
alignment htextalign = left!
string powertiptext = "Conversion settings"
end type

event clicked;OpenWithParm(w_dw2excelapp_settings, istr_currentsettings)

if not IsValid(Message.PowerObjectParm) then return

istr_currentsettings = Message.PowerObjectParm

of_savesettings(istr_currentsettings)
end event

type ddlb_column from dropdownlistbox within u_dwrenderer
integer x = 1111
integer y = 96
integer width = 626
integer height = 364
integer taborder = 20
boolean bringtotop = true
integer textsize = -10
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
long textcolor = 33554432
boolean sorted = false
boolean vscrollbar = true
borderstyle borderstyle = stylelowered!
end type

type hpb_exportprogress from hprogressbar within u_dwrenderer
integer y = 1804
integer width = 3296
integer height = 68
unsignedinteger maxposition = 100
integer setstep = 10
end type

type dw_host from datawindow within u_dwrenderer
integer y = 216
integer width = 3296
integer height = 1580
integer taborder = 20
string title = "none"
boolean hscrollbar = true
boolean vscrollbar = true
boolean livescroll = true
borderstyle borderstyle = stylelowered!
end type

type st_t_dwstyle from statictext within u_dwrenderer
integer y = 28
integer width = 480
integer height = 64
integer textsize = -10
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
long textcolor = 33554432
long backcolor = 553648127
string text = "DataWindow style"
boolean focusrectangle = false
end type

type cb_export from commandbutton within u_dwrenderer
integer x = 2871
integer y = 84
integer width = 425
integer height = 112
integer taborder = 10
integer textsize = -10
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
string text = "Save as XSLX"
end type

event clicked;of_exportxlsx(Parent)
end event

type st_t_column from statictext within u_dwrenderer
integer x = 1115
integer y = 28
integer width = 480
integer height = 64
integer textsize = -10
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
long textcolor = 33554432
long backcolor = 553648127
string text = "Column"
boolean focusrectangle = false
end type

type st_t_value from statictext within u_dwrenderer
integer x = 1947
integer y = 28
integer width = 480
integer height = 64
integer textsize = -10
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
long textcolor = 33554432
long backcolor = 553648127
string text = "Value"
boolean focusrectangle = false
end type

type cbx_usefilter from checkbox within u_dwrenderer
integer x = 686
integer y = 88
integer width = 402
integer height = 80
integer textsize = -10
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
long textcolor = 33554432
long backcolor = 553648127
string text = "Filter data"
end type

event clicked;parent.event ue_filterControlsVisible(this.Checked)

if not Checked then
	dw_host.SetFilter("")
	dw_host.Filter()
end if
end event

type ddlb_operator from dropdownlistbox within u_dwrenderer
integer x = 1760
integer y = 96
integer width = 165
integer height = 328
integer taborder = 10
boolean bringtotop = true
integer textsize = -10
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
long textcolor = 33554432
boolean sorted = false
boolean vscrollbar = true
string item[] = {"=","<",">","<>"}
borderstyle borderstyle = stylelowered!
end type

type st_t_operator from statictext within u_dwrenderer
integer x = 1765
integer y = 28
integer width = 142
integer height = 64
integer textsize = -10
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
long textcolor = 33554432
long backcolor = 553648127
string text = "Op"
boolean focusrectangle = false
end type

type pb_filter from picturebutton within u_dwrenderer
integer x = 2354
integer y = 96
integer width = 110
integer height = 96
integer taborder = 10
integer textsize = -8
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
string picturename = "Search1!"
alignment htextalign = left!
end type

event clicked;int res

string ls_filterExpression

ls_filterExpression = ddlb_column.Text + " " + ddlb_operator.Text + " " + sle_value.Text;
res = dw_host.SetFilter(ls_filterExpression)
dw_host.Filter()

end event

type ddlb_dwstyle from dropdownlistbox within u_dwrenderer
integer y = 92
integer width = 626
integer height = 364
integer taborder = 20
boolean bringtotop = true
integer textsize = -10
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
long textcolor = 33554432
boolean sorted = false
boolean vscrollbar = true
string item[] = {"Tabular","Grid","Freeform"}
borderstyle borderstyle = stylelowered!
end type

event selectionchanged;int i

ii_selecteddwindex = index

dw_Host.DataObject = is_datawindows[index]
of_updateFilterdefinition( )

end event

type sle_value from singlelineedit within u_dwrenderer
integer x = 1943
integer y = 96
integer width = 393
integer height = 100
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

