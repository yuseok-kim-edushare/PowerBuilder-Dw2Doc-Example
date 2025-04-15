$PBExportHeader$w_savetotemplate.srw
forward
global type w_savetotemplate from window
end type
type st_s_preview from statictext within w_savetotemplate
end type
type ddlb_templates from dropdownlistbox within w_savetotemplate
end type
type p_templatepreview from picture within w_savetotemplate
end type
type cb_cancel from commandbutton within w_savetotemplate
end type
type cb_select from commandbutton within w_savetotemplate
end type
type plb_thumbnailslist from picturelistbox within w_savetotemplate
end type
type st_s_selecttemplate from statictext within w_savetotemplate
end type
type p_pictureframe from picture within w_savetotemplate
end type
end forward

global type w_savetotemplate from window
integer width = 2994
integer height = 1768
boolean titlebar = true
string title = "Word Template Selection"
boolean controlmenu = true
windowtype windowtype = response!
long backcolor = 67108864
string icon = "DosEditIcon!"
boolean center = true
event ue_error ( string as_title,  string astr_message )
st_s_preview st_s_preview
ddlb_templates ddlb_templates
p_templatepreview p_templatepreview
cb_cancel cb_cancel
cb_select cb_select
plb_thumbnailslist plb_thumbnailslist
st_s_selecttemplate st_s_selecttemplate
p_pictureframe p_pictureframe
end type
global w_savetotemplate w_savetotemplate

type variables
constant string TEMPLATE_PATH = "res\templates\"

n_cst_pathtools 			ino_PathTools // Path Tools  .NET Assembly
n_cst_directorytools 	ino_DirTools // Directory Tools  .NET Assembly
nvo_templateloader 	ino_TemplateLoader // Template Loader  .NET Assembly
dotnetobject 			ino_SelectedTemplate // Selected Template object  .NET Assembly
str_templateparams 	istr_params // Parameters of the selected template
dotnetobject 			ino_Templates[]
string 					is_TemplateNames[]
end variables

forward prototypes
public subroutine wf_loadthumbnails (dropdownlistbox ddlb_templatelist, readonly string as_templateindex)
public subroutine wf_setthumbnail (ref picture ap_control, picture ap_frame, ref dotnetobject ano_template)
end prototypes

event ue_error(string as_title, string astr_message);MessageBox(as_title, astr_message)
end event

public subroutine wf_loadthumbnails (dropdownlistbox ddlb_templatelist, readonly string as_templateindex);string 			ls_error
int 				li_ResultCode
dotnetobject 	lno_TemplateIndex

li_ResultCode = ino_TemplateLoader.loadtemplates( as_templateindex, ref lno_TemplateIndex, ref ls_error)

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
	if lno_template.TemplateType <> istr_params.i_templatetype then
		continue
	end if
	ino_Templates[UpperBound(ino_Templates) + 1] = lno_template
	is_TemplateNames[UpperBound(is_TemplateNames) + 1] = lno_template.Name
	ddlb_templatelist.AddItem(lno_template.Name)
	SetNull(lno_template)
next
end subroutine

public subroutine wf_setthumbnail (ref picture ap_control, picture ap_frame, ref dotnetobject ano_template);/// Calculates the dimensions the picture must have to neatly fit into the provided frame

int li_FrameWidthPx,&
	li_FrameHeightPx,&
	li_PreviewWidthPx,&
	li_PreviewHeightPx,&
	li_Padding = 50
	
li_FrameWidthPx = UnitsToPixels(ap_frame.Width, XUnitsToPixels!)
li_FrameHeightPx = UnitsToPixels(ap_frame.Height, YUnitsToPixels!)
li_PreviewWidthPx = ano_Template.ThumbnailWidth
li_PreviewHeightPx = ano_Template.ThumbnailHeight

double ldbl_scale
ldbl_scale = Min(double(li_FrameWidthPx)/double(li_PreviewWidthPx),&
					  double(li_FrameHeightPx )/ double(li_PreviewHeightPx))

SetRedraw(false)
ap_control.OriginalSize = false
ap_control.Width = PixelsToUnits(li_PreviewWidthPx * ldbl_scale, XPixelsToUnits!) - li_Padding
ap_control.Height = PixelsToUnits(li_PreviewHeightPx * ldbl_scale, YPixelsToUnits!) - li_Padding
ap_control.X = ap_frame.X + (ap_frame.width - ap_control.Width) / 2
ap_control.Y = ap_Frame.Y + (ap_frame.height - ap_control.Height) / 2
ap_control.PictureName 	= TEMPLATE_PATH + ano_template.Thumbnail
SetRedraw(true)
end subroutine

on w_savetotemplate.create
this.st_s_preview=create st_s_preview
this.ddlb_templates=create ddlb_templates
this.p_templatepreview=create p_templatepreview
this.cb_cancel=create cb_cancel
this.cb_select=create cb_select
this.plb_thumbnailslist=create plb_thumbnailslist
this.st_s_selecttemplate=create st_s_selecttemplate
this.p_pictureframe=create p_pictureframe
this.Control[]={this.st_s_preview,&
this.ddlb_templates,&
this.p_templatepreview,&
this.cb_cancel,&
this.cb_select,&
this.plb_thumbnailslist,&
this.st_s_selecttemplate,&
this.p_pictureframe}
end on

on w_savetotemplate.destroy
destroy(this.st_s_preview)
destroy(this.ddlb_templates)
destroy(this.p_templatepreview)
destroy(this.cb_cancel)
destroy(this.cb_select)
destroy(this.plb_thumbnailslist)
destroy(this.st_s_selecttemplate)
destroy(this.p_pictureframe)
end on

event open;ino_PathTools 			= create n_cst_pathtools
ino_DirTools 			= create n_cst_directorytools
ino_TemplateLoader 	= create nvo_templateLoader


ino_PathTools.of_createondemand( )
ino_DirTools.of_createondemand( )
ino_TemplateLoader.of_createOnDemand()

istr_params = Message.Powerobjectparm

wf_LoadThumbnails(ddlb_templates, TEMPLATE_PATH +  "templates.json")

ddlb_templates.SelectItem(3)
ddlb_templates.event SelectionChanged(3)

if istr_params.b_silent then
	cb_select.event clicked()
end if

end event

event close;destroy ino_PathTools 
destroy ino_DirTools 
destroy ino_TemplateLoader 
end event

type st_s_preview from statictext within w_savetotemplate
integer x = 37
integer y = 228
integer width = 846
integer height = 64
integer textsize = -10
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
long textcolor = 33554432
long backcolor = 67108864
string text = "Preview"
boolean focusrectangle = false
end type

type ddlb_templates from dropdownlistbox within w_savetotemplate
integer x = 37
integer y = 100
integer width = 1216
integer height = 400
integer taborder = 40
integer textsize = -10
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
long textcolor = 33554432
borderstyle borderstyle = stylelowered!
end type

event selectionchanged;ino_SelectedTemplate = ino_Templates[index]

wf_setthumbnail( p_templatepreview, p_pictureframe, ino_SelectedTemplate)
cb_select.Enabled = true
end event

type p_templatepreview from picture within w_savetotemplate
integer x = 37
integer y = 324
integer width = 2885
integer height = 1120
boolean focusrectangle = false
end type

type cb_cancel from commandbutton within w_savetotemplate
integer x = 2519
integer y = 1492
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

event clicked;Close(parent)
end event

type cb_select from commandbutton within w_savetotemplate
integer x = 2075
integer y = 1492
integer width = 402
integer height = 112
integer taborder = 30
integer textsize = -10
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
boolean enabled = false
string text = "OK"
boolean default = true
end type

event clicked;
string ls_newFile

CloseWithReturn(parent, ino_SelectedTemplate)

//ls_newFile = ino_PathTools.of_changefileextension( , /*string as_newext */)
//CloseWithReturn(Parent,)
end event

type plb_thumbnailslist from picturelistbox within w_savetotemplate
boolean visible = false
integer x = 32
integer y = 128
integer width = 2418
integer height = 516
integer taborder = 10
integer textsize = -10
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
long textcolor = 33554432
boolean vscrollbar = true
borderstyle borderstyle = stylelowered!
long picturemaskcolor = 536870912
end type

type st_s_selecttemplate from statictext within w_savetotemplate
integer x = 37
integer y = 20
integer width = 846
integer height = 64
integer textsize = -10
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
long textcolor = 33554432
long backcolor = 67108864
string text = "Select a template"
boolean focusrectangle = false
end type

type p_pictureframe from picture within w_savetotemplate
integer x = 37
integer y = 324
integer width = 2885
integer height = 1120
boolean border = true
boolean focusrectangle = false
end type

