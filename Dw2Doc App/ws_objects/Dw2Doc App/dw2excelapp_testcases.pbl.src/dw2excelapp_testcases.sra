$PBExportHeader$dw2excelapp_testcases.sra
$PBExportComments$Generated SDI Application Object
forward
global type dw2excelapp_testcases from application
end type
global transaction sqlca
global dynamicdescriptionarea sqlda
global dynamicstagingarea sqlsa
global error error
global message message
end forward

global type dw2excelapp_testcases from application
string appname = "dw2excelapp_testcases"
string themepath = ".\res\theme"
string themename = "Flat Design Blue"
boolean nativepdfvalid = false
boolean nativepdfincludecustomfont = false
string nativepdfappname = ""
long richtextedittype = 5
long richtexteditx64type = 5
long richtexteditversion = 3
string richtexteditkey = ""
string appicon = ""
string appruntimeversion = "22.2.0.3397"
boolean manualsession = false
boolean unsupportedapierror = false
boolean ultrafast = false
boolean bignoreservercertificate = false
uint ignoreservercertificate = 0
long webview2distribution = 0
boolean webview2checkx86 = false
boolean webview2checkx64 = false
string webview2url = "https://developer.microsoft.com/en-us/microsoft-edge/webview2/"
end type
global dw2excelapp_testcases dw2excelapp_testcases

on dw2excelapp_testcases.create
appname="dw2excelapp_testcases"
message=create message
sqlca=create transaction
sqlda=create dynamicdescriptionarea
sqlsa=create dynamicstagingarea
error=create error
end on

on dw2excelapp_testcases.destroy
destroy(sqlca)
destroy(sqlda)
destroy(sqlsa)
destroy(error)
destroy(message)
end on

event open;//*-----------------------------------------------------------------*/
//*    open:  Application Open Script
//*           1) Opens Main window
//*-----------------------------------------------------------------*/
Open ( w_dw2excelapp_testcases_main )
end event

