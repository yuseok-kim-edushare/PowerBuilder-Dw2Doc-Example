$PBExportHeader$u_dwtemplateexporter.sru
forward
global type u_dwtemplateexporter from u_resizableuserobject
end type
type dw_master from datawindow within u_dwtemplateexporter
end type
type st_s_personid from statictext within u_dwtemplateexporter
end type
type st_id from statictext within u_dwtemplateexporter
end type
type st_s_persontype from statictext within u_dwtemplateexporter
end type
type st_type from statictext within u_dwtemplateexporter
end type
type st_s_email from statictext within u_dwtemplateexporter
end type
type st_email from statictext within u_dwtemplateexporter
end type
type st_s_phone from statictext within u_dwtemplateexporter
end type
type st_phone from statictext within u_dwtemplateexporter
end type
type st_s_persontitle from statictext within u_dwtemplateexporter
end type
type st_title from statictext within u_dwtemplateexporter
end type
type st_s_firstname from statictext within u_dwtemplateexporter
end type
type st_firstname from statictext within u_dwtemplateexporter
end type
type st_s_middlename from statictext within u_dwtemplateexporter
end type
type st_middlename from statictext within u_dwtemplateexporter
end type
type st_s_lastname from statictext within u_dwtemplateexporter
end type
type st_lastname from statictext within u_dwtemplateexporter
end type
type st_s_suffix from statictext within u_dwtemplateexporter
end type
type st_suffix from statictext within u_dwtemplateexporter
end type
type st_s_customerdetails from statictext within u_dwtemplateexporter
end type
type st_s_addressline1 from statictext within u_dwtemplateexporter
end type
type st_addressline1 from statictext within u_dwtemplateexporter
end type
type st_s_addressline2 from statictext within u_dwtemplateexporter
end type
type st_addressline2 from statictext within u_dwtemplateexporter
end type
type st_s_postalcode from statictext within u_dwtemplateexporter
end type
type st_postalcode from statictext within u_dwtemplateexporter
end type
type st_s_city from statictext within u_dwtemplateexporter
end type
type st_city from statictext within u_dwtemplateexporter
end type
type st_s_state from statictext within u_dwtemplateexporter
end type
type st_state from statictext within u_dwtemplateexporter
end type
type st_s_country from statictext within u_dwtemplateexporter
end type
type st_country from statictext within u_dwtemplateexporter
end type
type st_s_addressdetails from statictext within u_dwtemplateexporter
end type
type cbx_usedocxtemplate from checkbox within u_dwtemplateexporter
end type
type cb_saveword from commandbutton within u_dwtemplateexporter
end type
type st_selecteddocxtemplate from statictext within u_dwtemplateexporter
end type
type ddlb_docxtemplate from dropdownlistbox within u_dwtemplateexporter
end type
type gb_details from groupbox within u_dwtemplateexporter
end type
end forward

global type u_dwtemplateexporter from u_resizableuserobject
integer width = 3342
integer height = 1920
event ue_error ( string as_title,  string as_message )
dw_master dw_master
st_s_personid st_s_personid
st_id st_id
st_s_persontype st_s_persontype
st_type st_type
st_s_email st_s_email
st_email st_email
st_s_phone st_s_phone
st_phone st_phone
st_s_persontitle st_s_persontitle
st_title st_title
st_s_firstname st_s_firstname
st_firstname st_firstname
st_s_middlename st_s_middlename
st_middlename st_middlename
st_s_lastname st_s_lastname
st_lastname st_lastname
st_s_suffix st_s_suffix
st_suffix st_suffix
st_s_customerdetails st_s_customerdetails
st_s_addressline1 st_s_addressline1
st_addressline1 st_addressline1
st_s_addressline2 st_s_addressline2
st_addressline2 st_addressline2
st_s_postalcode st_s_postalcode
st_postalcode st_postalcode
st_s_city st_s_city
st_city st_city
st_s_state st_s_state
st_state st_state
st_s_country st_s_country
st_country st_country
st_s_addressdetails st_s_addressdetails
cbx_usedocxtemplate cbx_usedocxtemplate
cb_saveword cb_saveword
st_selecteddocxtemplate st_selecteddocxtemplate
ddlb_docxtemplate ddlb_docxtemplate
gb_details gb_details
end type
global u_dwtemplateexporter u_dwtemplateexporter

type variables
constant string TEMPLATE_PATH = "res\templates\"



nvo_docxwriter 					ino_DocxWriter // DOCX Writer  .NET Assembly
nvo_employeeaddressloader 	ino_AddressLoader // Address Loader  .NET Assembly
dotnetobject 						ino_AddressDetails // Address Details .NET Object
dotnetobject 						ino_AddressRepo // Address Repository  .NET Object
dotnetobject 						ino_DocxTemplates [] // Loaded DOCX Templates
n_cst_pathtools 						ino_PathTools // Path Tools  .NET Assembly
n_cst_directorytools 				ino_DirTools // Directory Tools  .NET Assembly
nvo_templateloader 				ino_TemplateLoader // Template Loader  .NET Assembly

string is_selectedDocxTemplatePath // holds the DOCX template that was selected
end variables

forward prototypes
public function integer wf_createdocx (ref dotnetobject ano_document, ref string as_error)
public function integer wf_opendocx (string as_file, ref dotnetobject ano_document, ref string as_error)
public function integer wf_savedocument (ref dotnetobject ano_document, string as_path, string as_error)
public subroutine wf_loadrowdetail (long al_row)
public subroutine wf_selectrow (integer as_row, datawindow adw)
public subroutine of_configureresize ()
public subroutine wf_loadpersondata ()
public subroutine wf_loadtemplates (string as_templateindexfile, dropdownlistbox ddlb_docxlistbox)
public subroutine of_init ()
public subroutine of_resize (long al_newwidth, long al_newheight)
public function integer wf_writedocument (ref dotnetobject ano_document, string as_columns[], string as_values[], boolean ab_usingtemplate, ref string as_error)
public function integer wf_builddata (readonly datawindow adw_datawindow, ref string as_columns[], ref string as_values[])
end prototypes

event ue_error(string as_title, string as_message);MessageBox(as_title, as_message)
end event

public function integer wf_createdocx (ref dotnetobject ano_document, ref string as_error);int 					il_ResultCode
dotnetobject		lno_document

il_ResultCode = ino_DocxWriter.createdocument(ref lno_document, ref as_error)

if il_ResultCode <> 1 then 
	event ue_error("Create Document", as_error)
	return -1
end if

ano_document = lno_document
return 1
end function

public function integer wf_opendocx (string as_file, ref dotnetobject ano_document, ref string as_error);int il_ResultCode 

il_ResultCode = ino_DocxWriter.OpenDocument(as_file, ref ano_document, ref as_error)

if il_ResultCode <> 1 then 
	event ue_error("Open DOCX Document", as_error)
end if

return il_ResultCode


end function

public function integer wf_savedocument (ref dotnetobject ano_document, string as_path, string as_error);int il_ResultCode

il_ResultCode = ino_DocxWriter.SaveDocument(ref ano_document, as_path, ref as_error)

if il_ResultCode <> 1 then
	event ue_error("Save DOCX Document", as_error)
end if

return il_ResultCode
end function

public subroutine wf_loadrowdetail (long al_row);
long ll_id // Entity ID

ll_id = dw_master.object.businessentityid[al_row]
st_id.text 					= string(ll_id)
st_type.text 				= dw_master.object.persontype[al_row]
st_title.text 				= dw_master.object.title[al_row]
st_firstname.text 		= dw_master.object.firstname[al_row]
st_middlename.text 		= dw_master.object.middlename[al_row]
st_lastname.text 			= dw_master.object.lastname[al_row]
st_suffix.text 			= dw_master.object.suffix[al_row]
st_email.text 				= dw_master.object.emailaddress[al_row]
st_phone.text				= dw_master.object.phonenumber[al_row]

int 				il_ResultCode
string 			ls_error
dotnetobject 	lno_result

if IsNull(ino_AddressRepo) then
	return
end if

il_ResultCode = ino_AddressLoader.SearchById(long(ll_id),&
			ref ino_AddressRepo, &
			ref lno_result, &
			ref ls_error)

if il_ResultCode <> 1 then
	event ue_error("Load Address Detail", ls_error)
	return
end if

ino_AddressDetails = lno_result

st_addressline1.Text 	= lno_result.AddressLine1
st_addressline2.Text 	= lno_result.AddressLine2
st_city.Text 				= lno_result.City
st_postalcode.Text 		= lno_result.PostalCode
st_country.Text 			= lno_result.CountryRegionCode
st_State.Text 				= lno_result.StateName + "(" + lno_result.StateProvinceCode + ")"
end subroutine

public subroutine wf_selectrow (integer as_row, datawindow adw);adw.SelectRow(0, false)
if as_row > 0 then
	adw.SelectRow(as_row, true)
	wf_LoadRowDetail(as_row)
	cb_saveword.Enabled = true
	Enabled = true
end if
end subroutine

public subroutine of_configureresize ();
end subroutine

public subroutine wf_loadpersondata ();int il_ResultCode

il_ResultCode = f_loadjsonintodatawindow(&
			"data\Person.json",&
			dw_master)

if il_ResultCode <= 0 then
	MessageBox("ImportJson", "Could not import JSON: " + string(il_ResultCode))
end if
end subroutine

public subroutine wf_loadtemplates (string as_templateindexfile, dropdownlistbox ddlb_docxlistbox);string 			ls_error
int 				li_ResultCode
dotnetobject 	lno_TemplateIndex

li_ResultCode = ino_TemplateLoader.loadtemplates( as_templateindexFile, ref lno_TemplateIndex, ref ls_error)

if li_ResultCode <> 1 then
	event ue_error( "Load template index", ls_error)
	return
end if

int 				i
dotnetobject 	lno_template

for i = 1 to ino_TemplateLoader.GetTemplateCount(ref lno_TemplateIndex)
	li_ResultCode = ino_TemplateLoader.GetTemplate(ref lno_TemplateIndex, i - 1, ref lno_template, ref ls_error)
	if li_ResultCode <> 1 then
		event ue_error("Read template index", ls_error)
		return
	end if
	choose case lno_template.TemplateType
		case 1
			// xlsx
		case 2
			ddlb_docxListBox.AddItem(lno_template.Name)
			ino_docxtemplates[UpperBound(ino_docxtemplates) + 1] = lno_template
			// docx
	end choose

	SetNull(lno_template)
next
end subroutine

public subroutine of_init ();destroy ino_DocxWriter
destroy ino_DirTools
destroy ino_PathTools
destroy ino_AddressLoader
destroy ino_templateloader

ino_DocxWriter 	= create nvo_docxwriter
ino_PathTools 		= create n_cst_pathtools
ino_DirTools 		= create n_cst_directorytools
ino_AddressLoader = create nvo_employeeaddressloader

ino_templateLoader = create nvo_templateloader

of_configureresize( )

ino_DocxWriter.of_createondemand()
ino_DirTools.of_createondemand()
ino_PathTools.of_createOnDemand()
ino_AddressLoader.of_createOnDemand()
ino_templateLoader.of_createOnDemand()

ddlb_docxtemplate.Enabled = false

int il_ResultCode
string ls_error
il_ResultCode = ino_AddressLoader.load("data\PersonAddress.json", ref ino_AddressRepo, ref ls_error)

if il_ResultCode <> 1 then 
	event ue_error("Load PersonAddress", "Could not load PersonAddress data: " + ls_error)
	SetNull(ino_AddressRepo)
end if

wf_LoadPersonData()
dw_master.SelectRow(1, true)
wf_SelectRow(1, dw_master)


SetNull(is_selecteddocxtemplatepath)

wf_loadtemplates(TEMPLATE_PATH + "\templates.json", ddlb_docxtemplate)

is_selectedDocxTemplatePath = ino_DocxTemplates[3].TemplateFile
st_selectedDocxTemplate.Text = ino_DocxTemplates[3].Name

cbx_usedocxtemplate.Checked = true

end subroutine

public subroutine of_resize (long al_newwidth, long al_newheight);long ll_targetwidth
long ll_targetheight
long ll_minwidth = 3342
long ll_minheight = 1920


ll_targetwidth = al_newwidth
if al_newwidth < ll_minwidth then 
	ll_targetwidth = ll_minwidth
end if

ll_targetheight = al_newheight
if al_newheight < ll_minheight then
	ll_targetheight = ll_minheight
end if

dw_master.height = ll_targetheight - 64
dw_master.width = ll_targetwidth - 1400

gb_details.x = dw_master.width + dw_master.x + 32
gb_details.height = dw_master.height + 32

st_s_customerdetails.x = gb_details.x + 64
st_s_personid.x = st_s_customerdetails.x
st_s_persontype.x = st_s_customerdetails.x
st_s_personid.x = st_s_customerdetails.x
st_s_email.x = st_s_customerdetails.x
st_s_phone.x = st_s_customerdetails.x
st_s_persontitle.x = st_s_customerdetails.x
st_s_lastname.x = st_s_customerdetails.x
st_s_firstname.x = st_s_customerdetails.x
st_s_middlename.x = st_s_customerdetails.x
st_s_suffix.x = st_s_customerdetails.x

st_id.x = st_s_customerdetails.x + st_s_customerdetails.width + 64
st_type.x = st_s_customerdetails.x + st_s_customerdetails.width + 64
st_email.x = st_s_customerdetails.x + st_s_customerdetails.width + 64
st_phone.x = st_s_customerdetails.x + st_s_customerdetails.width + 64
st_title.x = st_s_customerdetails.x + st_s_customerdetails.width + 64
st_lastname.x = st_s_customerdetails.x + st_s_customerdetails.width + 64
st_firstname.x = st_s_customerdetails.x + st_s_customerdetails.width + 64
st_middlename.x = st_s_customerdetails.x + st_s_customerdetails.width + 64
st_suffix.x = st_s_customerdetails.x + st_s_customerdetails.width + 64

st_s_addressdetails.x = st_s_customerdetails.x
st_s_addressline1.x = st_s_customerdetails.x
st_s_addressline2.x = st_s_customerdetails.x
st_s_city.x = st_s_customerdetails.x
st_s_state.x = st_s_customerdetails.x
st_s_country.x = st_s_customerdetails.x
st_s_postalcode.x = st_s_customerdetails.x

st_addressline1.x = st_s_customerdetails.x  + st_s_customerdetails.width + 64
st_addressline2.x = st_s_customerdetails.x  + st_s_customerdetails.width + 64
st_city.x = st_s_customerdetails.x  + st_s_customerdetails.width + 64
st_state.x = st_s_customerdetails.x  + st_s_customerdetails.width + 64
st_country.x = st_s_customerdetails.x  + st_s_customerdetails.width + 64
st_postalcode.x = st_s_customerdetails.x  + st_s_customerdetails.width + 64

cb_saveword.x = st_s_customerdetails.x
cbx_usedocxtemplate.x = st_s_customerdetails.x
end subroutine

public function integer wf_writedocument (ref dotnetobject ano_document, string as_columns[], string as_values[], boolean ab_usingtemplate, ref string as_error);int 		il_ResultCode
int 		i	
string 	ls_coltype
string 	ls_value

for i = 1 to UpperBound(as_columns) 
	il_ResultCode = ino_DocxWriter.WriteToDocument(&
				ref ano_document,&
				as_columns[i],&
				as_values[i],&
				"N/A",&
				not ab_usingTemplate,&
				ref as_error)
next

if il_ResultCode <> 1 then
	event ue_error("Write to DOCX", as_error)
end if

return il_ResultCode

end function

public function integer wf_builddata (readonly datawindow adw_datawindow, ref string as_columns[], ref string as_values[]);int 		li_columns
int 		il_ResultCode
string 	currentRow

li_columns = integer(adw_datawindow.Describe("DataWindow.Column.Count"))

int i
for i = 1 to li_columns 
	as_columns[i] = adw_datawindow.Describe("#" + string(i) + ".Name")
next

string ls_coltype
string ls_value
for i = 1 to li_columns 
	ls_coltype = adw_datawindow.Describe(as_columns[i]+".ColType")
	if ls_coltype = "int" or ls_coltype = "long" or ls_coltype = "number" then
		ls_value = string(adw_datawindow.GetItemNumber(1, as_columns[i]))
	else 
		ls_value = adw_datawindow.GetItemString(1, as_columns[i])
	end if

	if IsNull(ls_value) then
		ls_value = ""
	end if
	
	as_values[UpperBound(as_values) + 1] = ls_value
next

as_columns[UpperBound(as_columns) + 1] = "addressline1" 
as_columns[UpperBound(as_columns) + 1] = "addressline2" 
as_columns[UpperBound(as_columns) + 1] = "postalcode" 
as_columns[UpperBound(as_columns) + 1] = "city" 
as_columns[UpperBound(as_columns) + 1] = "state" 
as_columns[UpperBound(as_columns) + 1] = "country" 

as_values[UpperBound(as_values) + 1] = ino_AddressDetails.AddressLine1
as_values[UpperBound(as_values) + 1] = ino_AddressDetails.AddressLine2
as_values[UpperBound(as_values) + 1] = ino_AddressDetails.PostalCode
as_values[UpperBound(as_values) + 1] = ino_AddressDetails.City
as_values[UpperBound(as_values) + 1] = ino_AddressDetails.StateName + " (" + ino_AddressDetails.StateProvinceCode + ")"
as_values[UpperBound(as_values) + 1] = ino_AddressDetails.CountryRegionCode

return 1
end function

on u_dwtemplateexporter.create
int iCurrent
call super::create
this.dw_master=create dw_master
this.st_s_personid=create st_s_personid
this.st_id=create st_id
this.st_s_persontype=create st_s_persontype
this.st_type=create st_type
this.st_s_email=create st_s_email
this.st_email=create st_email
this.st_s_phone=create st_s_phone
this.st_phone=create st_phone
this.st_s_persontitle=create st_s_persontitle
this.st_title=create st_title
this.st_s_firstname=create st_s_firstname
this.st_firstname=create st_firstname
this.st_s_middlename=create st_s_middlename
this.st_middlename=create st_middlename
this.st_s_lastname=create st_s_lastname
this.st_lastname=create st_lastname
this.st_s_suffix=create st_s_suffix
this.st_suffix=create st_suffix
this.st_s_customerdetails=create st_s_customerdetails
this.st_s_addressline1=create st_s_addressline1
this.st_addressline1=create st_addressline1
this.st_s_addressline2=create st_s_addressline2
this.st_addressline2=create st_addressline2
this.st_s_postalcode=create st_s_postalcode
this.st_postalcode=create st_postalcode
this.st_s_city=create st_s_city
this.st_city=create st_city
this.st_s_state=create st_s_state
this.st_state=create st_state
this.st_s_country=create st_s_country
this.st_country=create st_country
this.st_s_addressdetails=create st_s_addressdetails
this.cbx_usedocxtemplate=create cbx_usedocxtemplate
this.cb_saveword=create cb_saveword
this.st_selecteddocxtemplate=create st_selecteddocxtemplate
this.ddlb_docxtemplate=create ddlb_docxtemplate
this.gb_details=create gb_details
iCurrent=UpperBound(this.Control)
this.Control[iCurrent+1]=this.dw_master
this.Control[iCurrent+2]=this.st_s_personid
this.Control[iCurrent+3]=this.st_id
this.Control[iCurrent+4]=this.st_s_persontype
this.Control[iCurrent+5]=this.st_type
this.Control[iCurrent+6]=this.st_s_email
this.Control[iCurrent+7]=this.st_email
this.Control[iCurrent+8]=this.st_s_phone
this.Control[iCurrent+9]=this.st_phone
this.Control[iCurrent+10]=this.st_s_persontitle
this.Control[iCurrent+11]=this.st_title
this.Control[iCurrent+12]=this.st_s_firstname
this.Control[iCurrent+13]=this.st_firstname
this.Control[iCurrent+14]=this.st_s_middlename
this.Control[iCurrent+15]=this.st_middlename
this.Control[iCurrent+16]=this.st_s_lastname
this.Control[iCurrent+17]=this.st_lastname
this.Control[iCurrent+18]=this.st_s_suffix
this.Control[iCurrent+19]=this.st_suffix
this.Control[iCurrent+20]=this.st_s_customerdetails
this.Control[iCurrent+21]=this.st_s_addressline1
this.Control[iCurrent+22]=this.st_addressline1
this.Control[iCurrent+23]=this.st_s_addressline2
this.Control[iCurrent+24]=this.st_addressline2
this.Control[iCurrent+25]=this.st_s_postalcode
this.Control[iCurrent+26]=this.st_postalcode
this.Control[iCurrent+27]=this.st_s_city
this.Control[iCurrent+28]=this.st_city
this.Control[iCurrent+29]=this.st_s_state
this.Control[iCurrent+30]=this.st_state
this.Control[iCurrent+31]=this.st_s_country
this.Control[iCurrent+32]=this.st_country
this.Control[iCurrent+33]=this.st_s_addressdetails
this.Control[iCurrent+34]=this.cbx_usedocxtemplate
this.Control[iCurrent+35]=this.cb_saveword
this.Control[iCurrent+36]=this.st_selecteddocxtemplate
this.Control[iCurrent+37]=this.ddlb_docxtemplate
this.Control[iCurrent+38]=this.gb_details
end on

on u_dwtemplateexporter.destroy
call super::destroy
destroy(this.dw_master)
destroy(this.st_s_personid)
destroy(this.st_id)
destroy(this.st_s_persontype)
destroy(this.st_type)
destroy(this.st_s_email)
destroy(this.st_email)
destroy(this.st_s_phone)
destroy(this.st_phone)
destroy(this.st_s_persontitle)
destroy(this.st_title)
destroy(this.st_s_firstname)
destroy(this.st_firstname)
destroy(this.st_s_middlename)
destroy(this.st_middlename)
destroy(this.st_s_lastname)
destroy(this.st_lastname)
destroy(this.st_s_suffix)
destroy(this.st_suffix)
destroy(this.st_s_customerdetails)
destroy(this.st_s_addressline1)
destroy(this.st_addressline1)
destroy(this.st_s_addressline2)
destroy(this.st_addressline2)
destroy(this.st_s_postalcode)
destroy(this.st_postalcode)
destroy(this.st_s_city)
destroy(this.st_city)
destroy(this.st_s_state)
destroy(this.st_state)
destroy(this.st_s_country)
destroy(this.st_country)
destroy(this.st_s_addressdetails)
destroy(this.cbx_usedocxtemplate)
destroy(this.cb_saveword)
destroy(this.st_selecteddocxtemplate)
destroy(this.ddlb_docxtemplate)
destroy(this.gb_details)
end on

event constructor;call super::constructor;of_init()
end event

type dw_master from datawindow within u_dwtemplateexporter
integer y = 32
integer width = 1943
integer height = 1872
integer taborder = 20
string title = "none"
string dataobject = "d_person"
boolean hscrollbar = true
boolean vscrollbar = true
boolean livescroll = true
borderstyle borderstyle = stylelowered!
end type

event clicked;wf_SelectRow(row, this)
end event

type st_s_personid from statictext within u_dwtemplateexporter
integer x = 2030
integer y = 204
integer width = 311
integer height = 64
integer textsize = -10
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
long textcolor = 33554432
long backcolor = 67108864
string text = "Id:"
boolean focusrectangle = false
end type

type st_id from statictext within u_dwtemplateexporter
integer x = 2578
integer y = 204
integer width = 640
integer height = 64
integer textsize = -10
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
long textcolor = 33554432
long backcolor = 67108864
string text = "Id:"
boolean focusrectangle = false
end type

type st_s_persontype from statictext within u_dwtemplateexporter
integer x = 2030
integer y = 284
integer width = 361
integer height = 64
integer textsize = -10
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
long textcolor = 33554432
long backcolor = 67108864
string text = "Person Type:"
boolean focusrectangle = false
end type

type st_type from statictext within u_dwtemplateexporter
integer x = 2578
integer y = 284
integer width = 640
integer height = 64
integer textsize = -10
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
long textcolor = 33554432
long backcolor = 67108864
string text = "Id:"
boolean focusrectangle = false
end type

type st_s_email from statictext within u_dwtemplateexporter
integer x = 2030
integer y = 364
integer width = 375
integer height = 64
integer textsize = -10
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
long textcolor = 33554432
long backcolor = 67108864
string text = "Email:"
boolean focusrectangle = false
end type

type st_email from statictext within u_dwtemplateexporter
integer x = 2578
integer y = 364
integer width = 640
integer height = 64
integer textsize = -10
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
long textcolor = 33554432
long backcolor = 67108864
string text = "Id:"
boolean focusrectangle = false
end type

type st_s_phone from statictext within u_dwtemplateexporter
integer x = 2030
integer y = 444
integer width = 375
integer height = 64
integer textsize = -10
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
long textcolor = 33554432
long backcolor = 67108864
string text = "Phone:"
boolean focusrectangle = false
end type

type st_phone from statictext within u_dwtemplateexporter
integer x = 2578
integer y = 444
integer width = 640
integer height = 64
integer textsize = -10
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
long textcolor = 33554432
long backcolor = 67108864
string text = "Id:"
boolean focusrectangle = false
end type

type st_s_persontitle from statictext within u_dwtemplateexporter
integer x = 2030
integer y = 524
integer width = 361
integer height = 64
integer textsize = -10
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
long textcolor = 33554432
long backcolor = 67108864
string text = "Title:"
boolean focusrectangle = false
end type

type st_title from statictext within u_dwtemplateexporter
integer x = 2578
integer y = 524
integer width = 640
integer height = 64
integer textsize = -10
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
long textcolor = 33554432
long backcolor = 67108864
string text = "Id:"
boolean focusrectangle = false
end type

type st_s_firstname from statictext within u_dwtemplateexporter
integer x = 2030
integer y = 604
integer width = 361
integer height = 64
integer textsize = -10
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
long textcolor = 33554432
long backcolor = 67108864
string text = "First Name:"
boolean focusrectangle = false
end type

type st_firstname from statictext within u_dwtemplateexporter
integer x = 2578
integer y = 604
integer width = 640
integer height = 64
integer textsize = -10
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
long textcolor = 33554432
long backcolor = 67108864
string text = "Id:"
boolean focusrectangle = false
end type

type st_s_middlename from statictext within u_dwtemplateexporter
integer x = 2030
integer y = 684
integer width = 375
integer height = 64
integer textsize = -10
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
long textcolor = 33554432
long backcolor = 67108864
string text = "Middle Name:"
boolean focusrectangle = false
end type

type st_middlename from statictext within u_dwtemplateexporter
integer x = 2578
integer y = 684
integer width = 640
integer height = 64
integer textsize = -10
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
long textcolor = 33554432
long backcolor = 67108864
string text = "Id:"
boolean focusrectangle = false
end type

type st_s_lastname from statictext within u_dwtemplateexporter
integer x = 2030
integer y = 764
integer width = 375
integer height = 64
integer textsize = -10
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
long textcolor = 33554432
long backcolor = 67108864
string text = "Last Name:"
boolean focusrectangle = false
end type

type st_lastname from statictext within u_dwtemplateexporter
integer x = 2578
integer y = 764
integer width = 640
integer height = 64
integer textsize = -10
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
long textcolor = 33554432
long backcolor = 67108864
string text = "Id:"
boolean focusrectangle = false
end type

type st_s_suffix from statictext within u_dwtemplateexporter
integer x = 2030
integer y = 844
integer width = 375
integer height = 64
integer textsize = -10
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
long textcolor = 33554432
long backcolor = 67108864
string text = "Suffix:"
boolean focusrectangle = false
end type

type st_suffix from statictext within u_dwtemplateexporter
integer x = 2578
integer y = 844
integer width = 640
integer height = 64
integer textsize = -10
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
long textcolor = 33554432
long backcolor = 67108864
string text = "Id:"
boolean focusrectangle = false
end type

type st_s_customerdetails from statictext within u_dwtemplateexporter
integer x = 2021
integer y = 92
integer width = 517
integer height = 76
integer textsize = -11
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
long textcolor = 33554432
long backcolor = 67108864
string text = "Customer Details"
boolean focusrectangle = false
end type

type st_s_addressline1 from statictext within u_dwtemplateexporter
integer x = 2030
integer y = 1068
integer width = 421
integer height = 64
integer textsize = -10
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
long textcolor = 33554432
long backcolor = 67108864
string text = "Address Line 1:"
boolean focusrectangle = false
end type

type st_addressline1 from statictext within u_dwtemplateexporter
integer x = 2578
integer y = 1068
integer width = 640
integer height = 64
integer textsize = -10
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
long textcolor = 33554432
long backcolor = 67108864
string text = "Id:"
boolean focusrectangle = false
end type

type st_s_addressline2 from statictext within u_dwtemplateexporter
integer x = 2030
integer y = 1148
integer width = 421
integer height = 64
integer textsize = -10
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
long textcolor = 33554432
long backcolor = 67108864
string text = "Address Line 2:"
boolean focusrectangle = false
end type

type st_addressline2 from statictext within u_dwtemplateexporter
integer x = 2578
integer y = 1148
integer width = 640
integer height = 64
integer textsize = -10
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
long textcolor = 33554432
long backcolor = 67108864
string text = "Id:"
boolean focusrectangle = false
end type

type st_s_postalcode from statictext within u_dwtemplateexporter
integer x = 2030
integer y = 1228
integer width = 375
integer height = 64
integer textsize = -10
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
long textcolor = 33554432
long backcolor = 67108864
string text = "Postal Code:"
boolean focusrectangle = false
end type

type st_postalcode from statictext within u_dwtemplateexporter
integer x = 2578
integer y = 1228
integer width = 640
integer height = 64
integer textsize = -10
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
long textcolor = 33554432
long backcolor = 67108864
string text = "Id:"
boolean focusrectangle = false
end type

type st_s_city from statictext within u_dwtemplateexporter
integer x = 2030
integer y = 1308
integer width = 375
integer height = 64
integer textsize = -10
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
long textcolor = 33554432
long backcolor = 67108864
string text = "City:"
boolean focusrectangle = false
end type

type st_city from statictext within u_dwtemplateexporter
integer x = 2578
integer y = 1308
integer width = 640
integer height = 64
integer textsize = -10
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
long textcolor = 33554432
long backcolor = 67108864
string text = "Id:"
boolean focusrectangle = false
end type

type st_s_state from statictext within u_dwtemplateexporter
integer x = 2030
integer y = 1388
integer width = 416
integer height = 64
integer textsize = -10
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
long textcolor = 33554432
long backcolor = 67108864
string text = "State/Province:"
boolean focusrectangle = false
end type

type st_state from statictext within u_dwtemplateexporter
integer x = 2578
integer y = 1388
integer width = 640
integer height = 64
integer textsize = -10
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
long textcolor = 33554432
long backcolor = 67108864
string text = "Id:"
boolean focusrectangle = false
end type

type st_s_country from statictext within u_dwtemplateexporter
integer x = 2030
integer y = 1468
integer width = 361
integer height = 64
integer textsize = -10
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
long textcolor = 33554432
long backcolor = 67108864
string text = "Country:"
boolean focusrectangle = false
end type

type st_country from statictext within u_dwtemplateexporter
integer x = 2578
integer y = 1468
integer width = 640
integer height = 64
integer textsize = -10
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
long textcolor = 33554432
long backcolor = 67108864
string text = "Id:"
boolean focusrectangle = false
end type

type st_s_addressdetails from statictext within u_dwtemplateexporter
integer x = 2021
integer y = 972
integer width = 517
integer height = 76
integer textsize = -11
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
long textcolor = 33554432
long backcolor = 67108864
string text = "Address Details"
boolean focusrectangle = false
end type

type cbx_usedocxtemplate from checkbox within u_dwtemplateexporter
integer x = 2030
integer y = 1584
integer width = 443
integer height = 80
integer textsize = -10
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
long textcolor = 33554432
long backcolor = 67108864
string text = "Use Template"
end type

event clicked;str_templateparams 	lstr_params
dotnetobject 			lno_selectedTemplate

if Checked then
	lstr_params.i_templatetype = 2
	OpenWithParm(w_savetotemplate, lstr_params)
	

	if IsValid(Message.PowerObjectparm) then
		lno_selectedTemplate = Message.PowerObjectParm
		
		is_selectedDocxTemplatePath = lno_selectedTemplate.TemplateFile
		st_selectedDocxTemplate.Text = lno_selectedTemplate.Name
	else
		Checked = false
	end if
	
else
	st_selecteddocxtemplate.Text = ""
	SetNull(is_selectedDocxTemplatePath)
end if
end event

type cb_saveword from commandbutton within u_dwtemplateexporter
integer x = 2030
integer y = 1764
integer width = 507
integer height = 112
integer taborder = 10
integer textsize = -10
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
boolean enabled = false
string text = "Save as Word"
end type

event clicked;int il_ResultCode
string ls_pathname
string ls_filename
string ls_error
string ls_columns[]
string ls_values[]
boolean lb_usingTemplate
dotnetobject ldn_document

lb_usingTemplate = false

SetPointer(HourGlass!)

if IsNull(is_selectedDocxTemplatePath) then
	il_ResultCode = wf_CreateDocx(ldn_document, ls_error)
else 
	il_ResultCode = wf_OpenDocx(TEMPLATE_PATH + is_selectedDocxTemplatePath, ldn_document, ls_error)
	lb_usingTemplate = true
end if

if IsNull(ldn_document) or not IsValid(ldn_document) then
	event ue_error( "Could not open/create document", ls_error)
	return
end if

wf_builddata( dw_master, ls_columns, ls_values)

il_ResultCode = wf_WriteDocument(ldn_document, ls_columns, ls_values, lb_usingTemplate, ls_error)

il_ResultCode = GetFileSaveName("Save document", ls_pathname, ls_filename, "docx", "Word Document (*.docx), *.docx", '', 11347)
if il_ResultCode <> 1 then
	return
end if

il_ResultCode = wf_SaveDocument(ldn_document, ls_pathname, ref ls_error)

if il_ResultCode = 1 then
	MessageBox("Save document", "Document saved successfully")
else
	MessageBox("Save document", ls_error)
end if




end event

type st_selecteddocxtemplate from statictext within u_dwtemplateexporter
integer x = 2030
integer y = 1668
integer width = 718
integer height = 64
integer textsize = -10
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
long textcolor = 33554432
long backcolor = 67108864
boolean focusrectangle = false
end type

type ddlb_docxtemplate from dropdownlistbox within u_dwtemplateexporter
boolean visible = false
integer x = 4114
integer y = 1576
integer width = 681
integer height = 400
integer taborder = 20
integer textsize = -10
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
long textcolor = 33554432
boolean sorted = false
borderstyle borderstyle = stylelowered!
end type

event selectionchanged;is_selectedDocxTemplatePath = ino_docxTemplates[index].TemplateFile
cb_saveword.Enabled = true
end event

type gb_details from groupbox within u_dwtemplateexporter
integer x = 1966
integer width = 1317
integer height = 1904
integer taborder = 60
integer textsize = -10
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
long textcolor = 33554432
long backcolor = 67108864
end type

