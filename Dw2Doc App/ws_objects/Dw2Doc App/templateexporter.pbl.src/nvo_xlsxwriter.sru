$PBExportHeader$nvo_xlsxwriter.sru
forward
global type nvo_xlsxwriter from dotnetobject
end type
end forward

global type nvo_xlsxwriter from dotnetobject
end type
global nvo_xlsxwriter nvo_xlsxwriter

type variables

PUBLIC:
String is_assemblypath = "bin\Appeon.DotnetDemo.DocumentWriter.dll"
String is_classname = "Appeon.DotnetDemo.DocumentWriter.XlsxWriter"

/*      Error types       */
Constant Int SUCCESS        =  0 // No error since latest reset
Constant Int LOAD_FAILURE   = -1 // Failed to load assembly
Constant Int CREATE_FAILURE = -2 // Failed to create .NET object
Constant Int CALL_FAILURE   = -3 // Call to .NET function failed

/* Latest error -- Public reset via of_ResetError */
PRIVATEWRITE Long il_ErrorType   
PRIVATEWRITE Long il_ErrorNumber 
PRIVATEWRITE String is_ErrorText 

PRIVATE:
/*  .NET object creation */
Boolean ib_objectCreated
end variables

forward prototypes
public subroutine of_setassemblyerror (long al_errortype, string as_actiontext, long al_errornumber, string as_errortext)
public subroutine of_reseterror ()
public function boolean of_createondemand ()
public function long of_writexlsx(string as_path,string as_columns[],string as_rows[],string as_separator,ref string as_error)
public function long of_opentemplate(string as_templatepath,ref dotnetobject anv_workbook,ref string as_error)
public function long of_createworkbook(ref dotnetobject anv_workbook,ref string as_error)
public function long of_writestringarraytoworkbook(ref dotnetobject anv_workbook,string as_data[],string as_separator,ref string as_error)
public function long of_savedocument(ref dotnetobject anv_workbook,string as_filename,ref string as_error)
public function long of_writexlsx(string as_path,any as_table,ref string as_error)
end prototypes

public subroutine of_setassemblyerror (long al_errortype, string as_actiontext, long al_errornumber, string as_errortext);
//*----------------------------------------------------------------------------------------------*/
//* PRIVATE of_setAssemblyError
//* Sets error description for specified error condition report by an assembly function
//*
//* Error description layout
//* 		| <actionText> failed.<EOL>
//* 		| Error Number: <errorNumber><EOL>
//* 		| Error Text: <errorText> (*)
//*  (*): Line skipped when <ErrorText> is empty
//*----------------------------------------------------------------------------------------------*/

/*    Format description */
String ls_error
ls_error = as_actionText + " failed.~r~n"
ls_error += "Error Number: " + String(al_errorNumber) + "."
If Len(Trim(as_errorText)) > 0 Then
	ls_error += "~r~nError Text: " + as_errorText
End If

/*  Retain state in instance variables */
This.il_ErrorType = al_errorType
This.is_ErrorText = ls_error
This.il_ErrorNumber = al_errorNumber
end subroutine

public subroutine of_reseterror ();
//*--------------------------------------------*/
//* PUBLIC of_ResetError
//* Clears previously registered error
//*--------------------------------------------*/

This.il_ErrorType = This.SUCCESS
This.is_ErrorText = ''
This.il_ErrorNumber = 0
end subroutine

public function boolean of_createondemand ();
//*--------------------------------------------------------------*/
//*  PUBLIC   of_createOnDemand( )
//*  Return   True:  .NET object created
//*               False: Failed to create .NET object
//*  Loads .NET assembly and creates instance of .NET class.
//*  Uses .NET when loading .NET assembly.
//*  Signals error If an error occurs.
//*  Resets any prior error when load + create succeeds.
//*--------------------------------------------------------------*/

This.of_ResetError( )
If This.ib_objectCreated Then Return True // Already created => DONE

Long ll_status 
String ls_action

/* Load assembly using .NET */
ls_action = 'Load ' + This.is_AssemblyPath
DotNetAssembly lnv_assembly
lnv_assembly = Create DotNetAssembly
ll_status = lnv_assembly.LoadWithDotNet(This.is_AssemblyPath)

/* Abort when load fails */
If ll_status <> 1 Then
	This.of_SetAssemblyError(This.LOAD_FAILURE, ls_action, ll_status, lnv_assembly.ErrorText)
	Return False // Load failed => ABORT
End If

/*   Create .NET object */
ls_action = 'Create ' + This.is_ClassName
ll_status = lnv_assembly.CreateInstance(is_ClassName, This)

/* Abort when create fails */
If ll_status <> 1 Then
	This.of_SetAssemblyError(This.CREATE_FAILURE, ls_action, ll_status, lnv_assembly.ErrorText)
	Return False // Load failed => ABORT
End If

This.ib_objectCreated = True
Return True
end function

public function long of_writexlsx(string as_path,string as_columns[],string as_rows[],string as_separator,ref string as_error);
//*-----------------------------------------------------------------*/
//*  .NET function : WriteXlsx
//*   Argument:
//*              String as_path
//*              String as_columns[]
//*              String as_rows[]
//*              String as_separator
//*              String as_error
//*   Return : Long
//*-----------------------------------------------------------------*/
/* Store the Return value from dotnet function */
Long ll_result

/* Create .NET object */
If Not This.of_createOnDemand( ) Then
	SetNull(ll_result)
	Return ll_result
End If

/* Trigger the dotnet function */
ll_result = This.writexlsx(as_path,as_columns,as_rows,as_separator,ref as_error)
Return ll_result
end function

public function long of_opentemplate(string as_templatepath,ref dotnetobject anv_workbook,ref string as_error);
//*-----------------------------------------------------------------*/
//*  .NET function : OpenTemplate
//*   Argument:
//*              String as_templatepath
//*              Dotnetobject anv_workbook
//*              String as_error
//*   Return : Long
//*-----------------------------------------------------------------*/
/* Store the Return value from dotnet function */
Long ll_result

/* Create .NET object */
If Not This.of_createOnDemand( ) Then
	SetNull(ll_result)
	Return ll_result
End If

/* Trigger the dotnet function */
ll_result = This.opentemplate(as_templatepath,ref anv_workbook,ref as_error)
Return ll_result
end function

public function long of_createworkbook(ref dotnetobject anv_workbook,ref string as_error);
//*-----------------------------------------------------------------*/
//*  .NET function : CreateWorkbook
//*   Argument:
//*              Dotnetobject anv_workbook
//*              String as_error
//*   Return : Long
//*-----------------------------------------------------------------*/
/* Store the Return value from dotnet function */
Long ll_result

/* Create .NET object */
If Not This.of_createOnDemand( ) Then
	SetNull(ll_result)
	Return ll_result
End If

/* Trigger the dotnet function */
ll_result = This.createworkbook(ref anv_workbook,ref as_error)
Return ll_result
end function

public function long of_writestringarraytoworkbook(ref dotnetobject anv_workbook,string as_data[],string as_separator,ref string as_error);
//*-----------------------------------------------------------------*/
//*  .NET function : WriteStringArrayToWorkbook
//*   Argument:
//*              Dotnetobject anv_workbook
//*              String as_data[]
//*              String as_separator
//*              String as_error
//*   Return : Long
//*-----------------------------------------------------------------*/
/* Store the Return value from dotnet function */
Long ll_result

/* Create .NET object */
If Not This.of_createOnDemand( ) Then
	SetNull(ll_result)
	Return ll_result
End If

/* Trigger the dotnet function */
ll_result = This.writestringarraytoworkbook(ref anv_workbook,as_data,as_separator,ref as_error)
Return ll_result
end function

public function long of_savedocument(ref dotnetobject anv_workbook,string as_filename,ref string as_error);
//*-----------------------------------------------------------------*/
//*  .NET function : SaveDocument
//*   Argument:
//*              Dotnetobject anv_workbook
//*              String as_filename
//*              String as_error
//*   Return : Long
//*-----------------------------------------------------------------*/
/* Store the Return value from dotnet function */
Long ll_result

/* Create .NET object */
If Not This.of_createOnDemand( ) Then
	SetNull(ll_result)
	Return ll_result
End If

/* Trigger the dotnet function */
ll_result = This.savedocument(ref anv_workbook,as_filename,ref as_error)
Return ll_result
end function

public function long of_writexlsx(string as_path,any as_table,ref string as_error);
//*-----------------------------------------------------------------*/
//*  .NET function : WriteXlsx
//*   Argument:
//*              String as_path
//*              String as_table[,]
//*              String as_error
//*   Return : Long
//*-----------------------------------------------------------------*/
/* Store the Return value from dotnet function */
Long ll_result

/* Create .NET object */
If Not This.of_createOnDemand( ) Then
	SetNull(ll_result)
	Return ll_result
End If

/* Trigger the dotnet function */
ll_result = This.writexlsx(as_path,as_table,ref as_error)
Return ll_result
end function

on nvo_xlsxwriter.create
call super::create
TriggerEvent( this, "constructor" )
end on

on nvo_xlsxwriter.destroy
TriggerEvent( this, "destructor" )
call super::destroy
end on

