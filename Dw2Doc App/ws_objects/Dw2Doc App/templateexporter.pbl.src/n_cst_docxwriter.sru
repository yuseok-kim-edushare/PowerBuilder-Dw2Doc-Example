$PBExportHeader$n_cst_docxwriter.sru
forward
global type n_cst_docxwriter from dotnetobject
end type
end forward

global type n_cst_docxwriter from dotnetobject
end type
global n_cst_docxwriter n_cst_docxwriter

type variables

PUBLIC:
String is_assemblypath = "bin\Appeon.DotnetDemo.DocxWriter.dll"
String is_classname = "Appeon.DotnetDemo.DocumentWriter.DocxWriter"

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
public function long of_opendocument(string as_path,ref dotnetobject anv_document,ref string as_error)
public function long of_createdocument(ref dotnetobject anv_document,ref string as_error)
public function long of_writetodocument(ref dotnetobject anv_document,string as_columnname,string as_columnvalue,string as_nullstring,ref string as_error)
public function long of_writetodocument(ref dotnetobject anv_document,string as_data[],string as_separator,string as_nullstring,ref string as_error)
public function long of_savedocument(ref dotnetobject anv_document,string as_path,ref string as_error)
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

public function long of_opendocument(string as_path,ref dotnetobject anv_document,ref string as_error);
//*-----------------------------------------------------------------*/
//*  .NET function : OpenDocument
//*   Argument:
//*              String as_path
//*              Dotnetobject anv_document
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
ll_result = This.opendocument(as_path,ref anv_document,ref as_error)
Return ll_result
end function

public function long of_createdocument(ref dotnetobject anv_document,ref string as_error);
//*-----------------------------------------------------------------*/
//*  .NET function : CreateDocument
//*   Argument:
//*              Dotnetobject anv_document
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
ll_result = This.createdocument(ref anv_document,ref as_error)
Return ll_result
end function

public function long of_writetodocument(ref dotnetobject anv_document,string as_columnname,string as_columnvalue,string as_nullstring,ref string as_error);
//*-----------------------------------------------------------------*/
//*  .NET function : WriteToDocument
//*   Argument:
//*              Dotnetobject anv_document
//*              String as_columnname
//*              String as_columnvalue
//*              String as_nullstring
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
ll_result = This.writetodocument(ref anv_document,as_columnname,as_columnvalue,as_nullstring,ref as_error)
Return ll_result
end function

public function long of_writetodocument(ref dotnetobject anv_document,string as_data[],string as_separator,string as_nullstring,ref string as_error);
//*-----------------------------------------------------------------*/
//*  .NET function : WriteToDocument
//*   Argument:
//*              Dotnetobject anv_document
//*              String as_data[]
//*              String as_separator
//*              String as_nullstring
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
ll_result = This.writetodocument(ref anv_document,as_data,as_separator,as_nullstring,ref as_error)
Return ll_result
end function

public function long of_savedocument(ref dotnetobject anv_document,string as_path,ref string as_error);
//*-----------------------------------------------------------------*/
//*  .NET function : SaveDocument
//*   Argument:
//*              Dotnetobject anv_document
//*              String as_path
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
ll_result = This.savedocument(ref anv_document,as_path,ref as_error)
Return ll_result
end function

on n_cst_docxwriter.create
call super::create
TriggerEvent( this, "constructor" )
end on

on n_cst_docxwriter.destroy
TriggerEvent( this, "destructor" )
call super::destroy
end on

