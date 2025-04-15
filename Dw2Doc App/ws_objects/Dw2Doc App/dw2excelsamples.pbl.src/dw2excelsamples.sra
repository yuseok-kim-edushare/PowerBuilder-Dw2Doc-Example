$PBExportHeader$dw2excelsamples.sra
$PBExportComments$Generated SDI Application Object
forward
global type dw2excelsamples from application
end type
global transaction sqlca
global dynamicdescriptionarea sqlda
global dynamicstagingarea sqlsa
global error error
global message message
end forward

global type dw2excelsamples from application
string appname = "dw2excelsamples"
string appruntimeversion = "22.2.0.3289"
end type
global dw2excelsamples dw2excelsamples

on dw2excelsamples.create
appname = "dw2excelsamples"
message = create message
sqlca = create transaction
sqlda = create dynamicdescriptionarea
sqlsa = create dynamicstagingarea
error = create error
end on

on dw2excelsamples.destroy
destroy( sqlca )
destroy( sqlda )
destroy( sqlsa )
destroy( error )
destroy( message )
end on

event open;//*-----------------------------------------------------------------*/
//*    open:  Application Open Script
//*           1) Opens Main window
//*-----------------------------------------------------------------*/
Open ( w_dw2excelsamples_main )
end event

