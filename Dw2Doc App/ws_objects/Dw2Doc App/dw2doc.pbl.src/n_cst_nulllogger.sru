$PBExportHeader$n_cst_nulllogger.sru
forward
global type n_cst_nulllogger from n_cst_logger
end type
end forward

global type n_cst_nulllogger from n_cst_logger
end type
global n_cst_nulllogger n_cst_nulllogger

forward prototypes
public subroutine of_error (string as_log)
public subroutine of_log (string as_log)
public subroutine of_warning (string as_log)
public function integer of_init (string as_file)
end prototypes

public subroutine of_error (string as_log);// do nothing
end subroutine

public subroutine of_log (string as_log);// do nothing
end subroutine

public subroutine of_warning (string as_log);// do nothing
end subroutine

public function integer of_init (string as_file);ii_FileHandle = 0

return 0
end function

on n_cst_nulllogger.create
call super::create
end on

on n_cst_nulllogger.destroy
call super::destroy
end on

