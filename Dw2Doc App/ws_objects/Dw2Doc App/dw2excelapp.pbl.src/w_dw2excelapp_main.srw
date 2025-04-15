$PBExportHeader$w_dw2excelapp_main.srw
$PBExportComments$Generated SDI Main Window
forward
global type w_dw2excelapp_main from window
end type
type uo_renderer from u_dwrenderer within w_dw2excelapp_main
end type
type st_s_description from u_unthemedstatictext within w_dw2excelapp_main
end type
type st_s_subtitle from u_unthemedstatictext within w_dw2excelapp_main
end type
type st_s_title from u_unthemedstatictext within w_dw2excelapp_main
end type
end forward

global type w_dw2excelapp_main from window
integer width = 3451
integer height = 2356
boolean titlebar = true
string title = "DataWindow to Excel"
boolean controlmenu = true
boolean minbox = true
boolean maxbox = true
boolean resizable = true
long backcolor = 79416533
string icon = "DataWindowIcon1!"
boolean center = true
uo_renderer uo_renderer
st_s_description st_s_description
st_s_subtitle st_s_subtitle
st_s_title st_s_title
end type
global w_dw2excelapp_main w_dw2excelapp_main

on w_dw2excelapp_main.create
this.uo_renderer=create uo_renderer
this.st_s_description=create st_s_description
this.st_s_subtitle=create st_s_subtitle
this.st_s_title=create st_s_title
this.Control[]={this.uo_renderer,&
this.st_s_description,&
this.st_s_subtitle,&
this.st_s_title}
end on

on w_dw2excelapp_main.destroy
destroy(this.uo_renderer)
destroy(this.st_s_description)
destroy(this.st_s_subtitle)
destroy(this.st_s_title)
end on

event resize;int li_targetWidth
int li_targetHeight
int li_minWidth = 3415
int li_minHeight = 2252

li_targetWidth = newwidth
li_targetHeight = newheight

if newwidth < li_minwidth then li_targetWidth = li_minWidth
if newheight < li_minheight then li_targetHeight = li_minHeight

uo_renderer.width = li_targetWidth - uo_renderer.x - 64

uo_renderer.height = li_targetHeight - uo_renderer.y - 60

uo_renderer.of_resize(uo_renderer.width, uo_renderer.height)
end event

type uo_renderer from u_dwrenderer within w_dw2excelapp_main
event destroy ( )
integer x = 64
integer y = 312
integer taborder = 30
end type

on uo_renderer.destroy
call u_dwrenderer::destroy
end on

type st_s_description from u_unthemedstatictext within w_dw2excelapp_main
integer x = 64
integer y = 156
integer width = 2021
long backcolor = 553648127
string text = "Select a DataWindow type and click Save as Excel to export it to an XLSX file"
end type

type st_s_subtitle from u_unthemedstatictext within w_dw2excelapp_main
integer x = 1029
integer y = 60
integer width = 1911
long textcolor = 8421504
long backcolor = 553648127
string text = "(Convert a DataWindow to Excel preserving the DW~'s visual appearance)"
end type

type st_s_title from u_unthemedstatictext within w_dw2excelapp_main
integer x = 64
integer y = 40
integer width = 1047
integer height = 96
integer textsize = -15
long textcolor = 134217741
long backcolor = 553648127
string text = "DataWindow to Excel"
end type

