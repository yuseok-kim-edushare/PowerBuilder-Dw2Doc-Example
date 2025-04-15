$PBExportHeader$w_dw2wordapp_main.srw
forward
global type w_dw2wordapp_main from window
end type
type uo_template from u_dwtemplateexporter within w_dw2wordapp_main
end type
type st_s_title from u_unthemedstatictext within w_dw2wordapp_main
end type
type st_s_subtitle from u_unthemedstatictext within w_dw2wordapp_main
end type
end forward

global type w_dw2wordapp_main from window
integer width = 3511
integer height = 2220
boolean titlebar = true
string title = "Dw2Word"
boolean controlmenu = true
boolean minbox = true
boolean maxbox = true
boolean resizable = true
long backcolor = 67108864
string icon = "DataWindowIcon!"
boolean center = true
uo_template uo_template
st_s_title st_s_title
st_s_subtitle st_s_subtitle
end type
global w_dw2wordapp_main w_dw2wordapp_main

on w_dw2wordapp_main.create
this.uo_template=create uo_template
this.st_s_title=create st_s_title
this.st_s_subtitle=create st_s_subtitle
this.Control[]={this.uo_template,&
this.st_s_title,&
this.st_s_subtitle}
end on

on w_dw2wordapp_main.destroy
destroy(this.uo_template)
destroy(this.st_s_title)
destroy(this.st_s_subtitle)
end on

event resize;int li_targetWidth
int li_targetHeight
int li_minWidth = 3474
int li_minHeight = 2116

li_targetWidth = newwidth
li_targetHeight = newheight

if newwidth < li_minwidth then li_targetWidth = li_minWidth
if newheight < li_minheight then li_targetHeight = li_minHeight

uo_template.width = li_targetWidth - uo_template.x - 64

uo_template.height = li_targetHeight - uo_template.y - 60

uo_template.of_resize(uo_template.width, uo_template.height)
end event

type uo_template from u_dwtemplateexporter within w_dw2wordapp_main
integer x = 64
integer y = 160
integer taborder = 30
end type

on uo_template.destroy
call u_dwtemplateexporter::destroy
end on

type st_s_title from u_unthemedstatictext within w_dw2wordapp_main
integer x = 64
integer y = 40
integer width = 1047
integer height = 96
integer textsize = -15
long textcolor = 134217741
long backcolor = 553648127
string text = "DataWindow to Word"
end type

type st_s_subtitle from u_unthemedstatictext within w_dw2wordapp_main
integer x = 1029
integer y = 60
integer width = 1911
long textcolor = 8421504
long backcolor = 553648127
string text = "(Convert a DataWindow to a Word Document using templates)"
end type

