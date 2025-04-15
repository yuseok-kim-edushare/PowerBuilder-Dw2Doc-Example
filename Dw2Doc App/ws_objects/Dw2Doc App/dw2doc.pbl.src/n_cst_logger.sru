$PBExportHeader$n_cst_logger.sru
forward
global type n_cst_logger from nonvisualobject
end type
end forward

global type n_cst_logger from nonvisualobject
end type
global n_cst_logger n_cst_logger

type variables
string is_path

protected int ii_FileHandle
end variables

forward prototypes
public subroutine of_log (string as_log)
public subroutine of_warning (string as_log)
public subroutine of_error (string as_log)
public function integer of_init (string as_file)
end prototypes

public subroutine of_log (string as_log);string ls_out

ls_out = "L[" + String(Today()) + " " + String(Now()) + "] " + as_log

FileWrite(ii_FileHandle, ls_out)
end subroutine

public subroutine of_warning (string as_log);string ls_out

ls_out = "W[" + String(Today()) + " " + String(Now()) + "] " + as_log

FileWrite(ii_FileHandle, ls_out)
end subroutine

public subroutine of_error (string as_log);string ls_out

ls_out = "E[" + String(Today()) + " " + String(Now()) + "] " + as_log

FileWrite(ii_FileHandle, ls_out)
end subroutine

public function integer of_init (string as_file);is_path = as_file

ii_FileHandle = FileOpen(is_path, LineMode!, Write!, LockWrite!, Append!, EncodingUTF8!)

return ii_fileHandle
end function

on n_cst_logger.create
call super::create
TriggerEvent( this, "constructor" )
end on

on n_cst_logger.destroy
TriggerEvent( this, "destructor" )
call super::destroy
end on

event destructor;if ii_FileHandle > 0 then FileClose(ii_FileHandle)
end event

