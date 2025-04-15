$PBExportHeader$w_dw2excelapp_testcases_main.srw
$PBExportComments$Generated SDI Main Window
forward
global type w_dw2excelapp_testcases_main from window
end type
type st_failurecount from statictext within w_dw2excelapp_testcases_main
end type
type cbx_onlyfailures from checkbox within w_dw2excelapp_testcases_main
end type
type st_s_results from statictext within w_dw2excelapp_testcases_main
end type
type hpb_progress from hprogressbar within w_dw2excelapp_testcases_main
end type
type dw_dwpreview from datawindow within w_dw2excelapp_testcases_main
end type
type ddlb_testcases from dropdownlistbox within w_dw2excelapp_testcases_main
end type
type st_s_testcases from statictext within w_dw2excelapp_testcases_main
end type
type cb_runselected from commandbutton within w_dw2excelapp_testcases_main
end type
type cb_runall from commandbutton within w_dw2excelapp_testcases_main
end type
type dw_testresults from datawindow within w_dw2excelapp_testcases_main
end type
end forward

global type w_dw2excelapp_testcases_main from window
integer width = 3314
integer height = 1920
boolean titlebar = true
string title = "Dw2Excel Test Cases"
boolean controlmenu = true
boolean minbox = true
boolean maxbox = true
boolean resizable = true
long backcolor = 79416533
string icon = "res\icons\flask.ico"
boolean center = true
event ue_conversionprogress ( long progress )
st_failurecount st_failurecount
cbx_onlyfailures cbx_onlyfailures
st_s_results st_s_results
hpb_progress hpb_progress
dw_dwpreview dw_dwpreview
ddlb_testcases ddlb_testcases
st_s_testcases st_s_testcases
cb_runselected cb_runselected
cb_runall cb_runall
dw_testresults dw_testresults
end type
global w_dw2excelapp_testcases_main w_dw2excelapp_testcases_main

type variables
n_cst_attributetesterlocator ino_locator
n_cst_dw2excelconverter ino_converter
n_cst_accessor ino_accessor
n_cst_stringextensions ino_string
string is_selectedDw
string is_dwnames[]
int ii_failureCount

end variables

forward prototypes
public subroutine of_setprogressbarvisible (boolean ab_visible)
public subroutine of_setprogress (integer ai_progress)
public subroutine of_testdw (datawindow adw_dw)
public subroutine of_deployresults (dotnetobject ano_results, string as_dwobject)
public subroutine of_loaddatawindows ()
public subroutine of_deployresult (string as_dwobject, string as_object, string as_attribute, string as_expected, string as_real, boolean ab_success)
end prototypes

event ue_conversionprogress(long progress);long ll_current
long ll_total

ll_current = Message.WordParm
ll_total = Message.LongParm


hpb_progress.Position = integer(100.0 * (double(ll_current) * double(ll_total)))
end event

public subroutine of_setprogressbarvisible (boolean ab_visible);hpb_progress.visible = ab_visible
end subroutine

public subroutine of_setprogress (integer ai_progress);int li_targetProgress

li_targetProgress = ai_progress
if ai_progress < 0 then li_targetProgress = 0
if ai_progress > 100 then li_targetProgress = 100

hpb_progress.Position = li_targetProgress
end subroutine

public subroutine of_testdw (datawindow adw_dw);dotnetobject lno_exportedcells
dotnetobject lno_exportedCell
dotnetobject lno_tester
dotnetobject ino_testresults
datastore lds
string ls_error
int i

lds = f_dw2ds(adw_dw, Primary!)

if not directoryExists("tests") then
	CreateDirectory("tests")
end if

try 
	SetNull(ls_error)
	
	ino_converter.of_convert(&
		lds,&
		"tests\" + adw_dw.dataobject + "_test.xlsx",&
		false,&
		"",&
		ls_error,&
		this,&
		"ue_conversionprogress"&
		)
		
	lno_exportedcells = ino_converter.ino_ConversionResult
	
	if IsNull(ls_error) then ls_error = ""
	
	of_deployresult( adw_dw.DataObject, "", "conversion error", "", ls_error, ls_error = "")	
	
	for i = 1 to lno_exportedCells.Count 
		lno_exportedCell = ino_accessor.Get(lno_exportedcells, i - 1)
		
		lno_tester = ino_locator.Find(lno_exportedCell)
		
		ino_testresults = lno_tester.Test(lno_exportedCell.Attribute, lno_exportedCell)
		
		of_deployresults(ino_testresults, adw_dw.dataobject)
		
		do while Yield()
		loop
	next
catch (Exception e)
	MessageBox("Error testing datawindow " + adw_dw.dataobject, e.GetMessage())
catch (RuntimeError re)
	MessageBox("Error testing datawindow " + adw_dw.dataobject, re.GetMessage())
finally
	destroy  lno_exportedcells
	destroy  lno_exportedCell
	destroy  lno_tester
	destroy ino_testresults
	hpb_progress.Position = 0
end try
end subroutine

public subroutine of_deployresults (dotnetobject ano_results, string as_dwobject);int i
int row
dotnetobject lno_result
boolean lno_success

for i = 1 to ano_results.Results.Count
	lno_result = ano_results.Get(i - 1)
	of_deployresult(&
		as_dwobject,&
		lno_result.ObjectName,&
		lno_result.Attribute,&
		lno_result.ExpectedValue,&
		lno_result.RealValue,&
		lno_result.Result)
next
end subroutine

public subroutine of_loaddatawindows ();int i 
string ls_dws
string ls_objects[]
string ls_tokens[]

ddlb_testcases.Reset( )

ls_dws = LibraryDirectory("test_dws.pbl", DirDataWindow!)

ino_string.of_split(ls_dws, "~n", ref ls_objects)

for i = 1 to UpperBound(ls_objects) 
	if ls_objects[i] = "" then continue
	ino_string.of_split(ls_objects[i], "~t", ref ls_tokens)
	is_dwnames[UpperBound(is_dwnames) + 1] = ls_tokens[1]

next

for i = 1 to  UpperBound(is_dwnames)
	ddlb_testcases.Additem(is_dwnames[i])
next
end subroutine

public subroutine of_deployresult (string as_dwobject, string as_object, string as_attribute, string as_expected, string as_real, boolean ab_success);int row
boolean lno_success

row = dw_testresults.InsertRow(0)

dw_testresults.object.caseid[row] = row
dw_testresults.object.description[row] = as_attribute
dw_testresults.object.dwobject[row] = as_dwobject
dw_testresults.object.object[row] = as_object
dw_testresults.object.expectedresult[row] = as_expected
dw_testresults.object.actualresult[row] = as_real
if ab_success then 
	dw_testresults.object.success[row] = 1
else
	dw_testresults.object.success[row] = 0
	ii_failurecount++
	st_failurecount.Text = "Failures: " + string(ii_failureCount)
end if

dw_testresults.ScrollToRow(dw_testresults.rowcount() + 1)

end subroutine

on w_dw2excelapp_testcases_main.create
this.st_failurecount=create st_failurecount
this.cbx_onlyfailures=create cbx_onlyfailures
this.st_s_results=create st_s_results
this.hpb_progress=create hpb_progress
this.dw_dwpreview=create dw_dwpreview
this.ddlb_testcases=create ddlb_testcases
this.st_s_testcases=create st_s_testcases
this.cb_runselected=create cb_runselected
this.cb_runall=create cb_runall
this.dw_testresults=create dw_testresults
this.Control[]={this.st_failurecount,&
this.cbx_onlyfailures,&
this.st_s_results,&
this.hpb_progress,&
this.dw_dwpreview,&
this.ddlb_testcases,&
this.st_s_testcases,&
this.cb_runselected,&
this.cb_runall,&
this.dw_testresults}
end on

on w_dw2excelapp_testcases_main.destroy
destroy(this.st_failurecount)
destroy(this.cbx_onlyfailures)
destroy(this.st_s_results)
destroy(this.hpb_progress)
destroy(this.dw_dwpreview)
destroy(this.ddlb_testcases)
destroy(this.st_s_testcases)
destroy(this.cb_runselected)
destroy(this.cb_runall)
destroy(this.dw_testresults)
end on

event resize;long ll_minwidth = 2085
long ll_minheight = 1296

long ll_targetwidth
long ll_targetheight

ll_targetwidth = newwidth
if newwidth < ll_minwidth then ll_targetwidth = ll_minwidth

ll_targetheight = newheight
if newheight < ll_minheight then ll_targetheight = ll_minheight

cb_runall.x = ll_targetwidth - cb_runall.width - 64
cb_runselected.x = cb_runall.x - cb_runselected.width - 32

dw_dwpreview.width = ll_targetwidth - 128
dw_dwpreview.height = (ll_targetheight - dw_dwpreview.y) * 0.5 
hpb_progress.y = dw_dwpreview.height + dw_dwpreview.y
hpb_progress.width=  dw_dwpreview.width

st_s_results.y = hpb_progress.height + hpb_progress.y + 32
st_failurecount.y = st_s_results.y
cbx_onlyfailures.y = st_s_results.y - 8
dw_testresults.y = st_s_results.y + st_s_results.height + 16
dw_testresults.height = ll_targetheight - dw_testresults.y - 64
dw_testresults.width = dw_dwpreview.width
end event

event open;string ls_error

ino_accessor = create n_cst_accessor

if not ino_accessor.of_createondemand() then
	MessageBox("n_cst_accessor error", ino_accessor.is_errortext)
	
	return
end if

SetNull(ls_error)

ino_locator = create n_cst_attributetesterlocator

if not ino_locator.of_createondemand( ) then
	MessageBox("n_cst_attributetesterlocator error", ino_locator.is_errortext)
	
	return
end if

ino_converter = create n_cst_dw2excelconverter

ino_converter.of_init(ls_error) 

if not IsNull(ls_error) then
	MessageBox("n_cst_dw2excelconverter error", ls_error)
	
	return
end if

ino_string = create n_cst_stringextensions

if not ino_string.of_createondemand() then 
	MessageBox("n_cst_stringextensions error", ino_string.is_errortext)
	return
end if

of_loaddatawindows( )

ddlb_testcases.SelectItem( 1)
ddlb_testCases.event SelectionChanged(1)


end event

event close;destroy ino_locator
destroy ino_string
destroy ino_converter
destroy ino_accessor
end event

type st_failurecount from statictext within w_dw2excelapp_testcases_main
integer x = 594
integer y = 744
integer width = 398
integer height = 64
integer textsize = -10
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
long textcolor = 33554432
long backcolor = 67108864
alignment alignment = right!
boolean focusrectangle = false
end type

type cbx_onlyfailures from checkbox within w_dw2excelapp_testcases_main
integer x = 1010
integer y = 736
integer width = 562
integer height = 80
integer textsize = -10
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
long textcolor = 33554432
long backcolor = 67108864
string text = "Show only failures"
end type

event clicked;
if Checked then
	dw_testresults.SetFilter("success = 0")
	dw_testresults.Filter()
else
	dw_testresults.SetFilter("")
	dw_testresults.Filter()
end if

dw_testresults.SetSort("caseid a")
end event

type st_s_results from statictext within w_dw2excelapp_testcases_main
integer x = 64
integer y = 744
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
string text = "Test Results:"
boolean focusrectangle = false
end type

type hpb_progress from hprogressbar within w_dw2excelapp_testcases_main
integer x = 64
integer y = 640
integer width = 1975
integer height = 68
unsignedinteger maxposition = 100
integer setstep = 10
end type

type dw_dwpreview from datawindow within w_dw2excelapp_testcases_main
integer x = 64
integer y = 244
integer width = 1975
integer height = 400
integer taborder = 20
string title = "none"
boolean hscrollbar = true
boolean vscrollbar = true
boolean livescroll = true
borderstyle borderstyle = stylelowered!
end type

type ddlb_testcases from dropdownlistbox within w_dw2excelapp_testcases_main
integer x = 64
integer y = 136
integer width = 814
integer height = 532
integer taborder = 20
integer textsize = -8
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
long textcolor = 33554432
boolean sorted = false
string item[] = {"dw_grid"}
borderstyle borderstyle = stylelowered!
end type

event selectionchanged;is_selectedDw = text

dw_dwpreview.DataObject = is_selecteddw

cb_runselected.Enabled = true
end event

type st_s_testcases from statictext within w_dw2excelapp_testcases_main
integer x = 64
integer y = 44
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
string text = "Test Cases:"
boolean focusrectangle = false
end type

type cb_runselected from commandbutton within w_dw2excelapp_testcases_main
integer x = 1198
integer y = 116
integer width = 402
integer height = 112
integer taborder = 10
integer textsize = -10
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
boolean enabled = false
string text = "Run Selected"
end type

event clicked;dw_testresults.Reset()

ii_failurecount = 0
st_failurecount.Text = ""

of_testdw(dw_dwpreview)

MessageBox("Message", "Test run finished")
end event

type cb_runall from commandbutton within w_dw2excelapp_testcases_main
integer x = 1637
integer y = 116
integer width = 402
integer height = 112
integer taborder = 10
integer textsize = -10
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
string text = "Run All"
end type

event clicked;int i 
datawindow ldw_dw

ii_failurecount = 0
st_failurecount.Text = ""

dw_testresults.Reset()

for i = 1 to UpperBound(is_dwnames)
	dw_dwpreview.reset()
	dw_dwpreview.DataObject = is_dwnames[i]
	of_testdw(dw_dwpreview)
next

MessageBox("Tests finished", "All tests have completed")
end event

type dw_testresults from datawindow within w_dw2excelapp_testcases_main
integer x = 64
integer y = 816
integer width = 1975
integer height = 472
integer taborder = 10
string title = "none"
string dataobject = "d_testresults"
boolean hscrollbar = true
boolean vscrollbar = true
boolean livescroll = true
borderstyle borderstyle = stylelowered!
end type

