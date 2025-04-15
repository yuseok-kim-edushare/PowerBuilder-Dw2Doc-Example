$PBExportHeader$u_resizableuserobject.sru
forward
global type u_resizableuserobject from userobject
end type
end forward

global type u_resizableuserobject from userobject
integer width = 503
integer height = 864
long backcolor = 67108864
string text = "none"
long tabtextcolor = 33554432
long picturemaskcolor = 536870912
end type
global u_resizableuserobject u_resizableuserobject

forward prototypes
public subroutine of_resize (long al_newwidth, long al_newheight)
end prototypes

public subroutine of_resize (long al_newwidth, long al_newheight);
end subroutine

on u_resizableuserobject.create
end on

on u_resizableuserobject.destroy
end on

