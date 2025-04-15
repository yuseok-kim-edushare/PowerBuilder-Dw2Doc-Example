$PBExportHeader$n_cst_dw2excelconverter.sru
forward
global type n_cst_dw2excelconverter from nonvisualobject
end type
end forward

global type n_cst_dw2excelconverter from nonvisualobject
end type
global n_cst_dw2excelconverter n_cst_dw2excelconverter

type variables
public:
n_cst_stringextensions ino_StringExtensions
n_cst_objectextensions ino_ObjectExtensions
n_cst_dwtools ino_dwTools
dotnetobject ino_ConversionResult
str_renderingsettings istr_settings

private:
n_cst_bitoperations ino_BitOperations
n_cst_colorfactory ino_ColorFactory
n_cst_codetabletools ino_CodeTableTools
n_cst_logger ino_logger
string is_bandstoinclude[] = {"header", "summary", "footer"}


end variables

forward prototypes
public function integer of_init (ref string as_error)
private function dotnetobject of_color2dwcolor (long al_color)
private function dotnetobject of_dwcontroltoattribute (datastore ads_source, readonly long al_index, readonly string as_controlname, readonly n_cst_dwattributefactory anv_dwattributefactory, ref string as_error)
private function string of_evaluateexpression (datastore ads_source, string as_expression, long al_row)
private function boolean of_getcontrolformat (datastore ads_source, string as_control, long al_row, ref string as_formatstring)
private function string of_getvalueforcurrentrow (datastore ads_source, string as_control, string as_property, long al_row)
public function boolean of_getpropertyexpression (datastore ads_source, string as_control, string as_property, ref string as_value, ref string as_expression)
public function integer of_convert (readonly datastore ads_ds, readonly string as_path)
public function integer of_convert (datastore ads_ds, string as_path, boolean ab_append, string as_sheetname, ref string as_error, powerobject ao_callback, string as_callbackevent)
public function integer of_convert (readonly datawindow adw_dw, readonly string as_path)
public function integer of_convert (readonly datawindow adw_dw, readonly string as_path, boolean ab_append, readonly string as_sheetname, ref string as_error, readonly powerobject ao_callback, readonly string as_callbackevent)
end prototypes

public function integer of_init (ref string as_error);ino_stringextensions = create n_cst_stringextensions
ino_stringextensions.of_createOnDemand()

if ino_stringextensions.il_errornumber <> 0 then
	as_error = "StringExtensions:" + ino_stringextensions.is_errorText
	return -1
end if

ino_dwTools = create n_cst_dwtools
ino_dwTools.of_CreateOnDemand()

if ino_dwTools.il_errornumber <> 0 then
	as_error = "DwTools" + ino_dwTools.is_errorText
	return -1
end if

ino_ObjectExtensions  = create n_cst_ObjectExtensions
ino_ObjectExtensions.of_createOnDemand()

if ino_ObjectExtensions.il_errornumber <> 0 then
	as_error = "ObjectExtensions:" + ino_ObjectExtensions.is_errorText
	return -1
end if

ino_ColorFactory = create n_cst_ColorFactory
ino_ColorFactory.of_createOnDemand()

if ino_ColorFactory.il_errornumber <> 0 then
	as_error = "ColorFactory:" + ino_ColorFactory.is_errorText
	return -1
end if

ino_BitOperations = create n_cst_BitOperations
ino_BitOperations.of_createOnDemand()

if ino_BitOperations.il_errornumber <> 0 then
	as_error = "BitOperations:" + ino_BitOperations.is_errorText
	return -1
end if

ino_CodeTableTools = create n_cst_CodeTableTools
ino_CodeTableTools.of_createOnDemand()

if ino_CodeTableTools.il_errornumber <> 0 then
	as_error = "CodetableTools: "  + ino_CodeTableTools.is_errorText
	return -1
end if

return 0

end function

private function dotnetobject of_color2dwcolor (long al_color);byte lby_red, lby_green, lby_blue, lby_alpha
boolean lb_isTransparent
string ls_colorname
long ll_systemcolor

lb_isTransparent = this.ino_BitOperations.BitwiseAndInt(al_color, this.ino_BitOperations.ShiftLeftInt(2, 28)) > 0

if lb_isTransparent then
	lby_alpha = 0
else
	lby_alpha = 255 
end if

if al_color > 16777215 then // Color value is not RGB
	if f_getsystemcolor(al_color, ref ll_systemcolor, ref ls_colorname) then 
		al_color = ll_systemcolor
	elseif not lb_isTransparent then
		ino_logger.of_warning("Unsupported color: " + ls_colorname)
	end if
end if

lby_red = byte(this.ino_BitOperations.BitwiseAndInt(al_color, 255))
lby_green = byte(this.ino_BitOperations.BitwiseAndInt(this.ino_BitOperations.ShiftRightInt(al_color, 8), 255))
lby_blue = byte(this.ino_BitOperations.BitwiseAndInt(this.ino_BitOperations.ShiftRightInt(al_color, 16), 255))

return this.ino_ColorFactory.FromArgb(lby_alpha, lby_red, lby_green, lby_blue)
end function

private function dotnetobject of_dwcontroltoattribute (datastore ads_source, readonly long al_index, readonly string as_controlname, readonly n_cst_dwattributefactory anv_dwattributefactory, ref string as_error);dotnetobject lnv_attribute
dotnetobject lnv_codetable
string ls_controltype
string ls_format
string ls_tmp
string ls_picturename
byte lby_shapeType
string ls_coltype
string ls_outputText
string ls_editstyle
string ls_unformattedValue
string ls_error
string ls_codetabledisplay[]
string ls_codetabledata[]
string ls_datatypeflag

ls_controltype = ads_source.Describe(as_controlname + ".Type")
if as_controlname = "t_7" then
	lby_shapeType = 0
end if
choose case ls_controltype
	case "text"
		lnv_attribute = anv_dwattributefactory.CreateTextAttributes()
		ls_outputText = of_GetValueForCurrentRow(ads_source, as_controlName, "Text", al_index)
		if left(ls_outputText, 1) = '"' and right(ls_outputText, 1) = '"' then
			ls_outputText = Mid(ls_outputText, 2, Len(ls_outputText) - 2)
		end if
		
		goto text_common
	case "column"
		ls_picturename = ads_source.Describe(as_controlname + ".BitmapName")
		
		if ls_picturename = "yes" then
			lnv_attribute = anv_dwAttributeFactory.CreatePictureAttributes()
			lnv_attribute.IsVisible = (of_GetValueForCurrentRow(ads_source, as_controlName, "Visible", al_index)) = "1"				
			lnv_attribute.FileName = ads_source.GetItemString(al_index, as_controlname)/*picture name*/
			lnv_attribute.Transparency = byte(of_GetValueForCurrentRow(ads_source, as_controlName, "Transparency", al_index))

			goto endchoose
		end if
		
case_column:
		
		ls_editstyle = ads_source.Describe(as_controlname + ".Edit.Style") 
		SetNull(as_error)

		choose case ls_editstyle
			case "radiobuttons"
				lnv_attribute = anv_dwattributefactory.CreateRadioButtonAttribute()
				
				ls_tmp = ads_source.Describe(as_controlname + ".Values")
				
				this.ino_codetabletools.GetValueMap(&
						ls_tmp,&
						ref lnv_codetable,&
						ref ls_error)
						
				if not IsNull(ls_error) then
					as_error = "Could not parse codetable for object " + as_controlname + ". " + ls_error
					return lnv_attribute
				end if
				
				ls_tmp = ads_source.Describe(as_controlname + ".radiobuttons.columns")
				
				lnv_attribute.Columns = integer(ls_tmp)
				lnv_attribute.LeftText = ads_source.Describe(as_controlname + ".checkbox.lefttext") = "yes"
				lnv_attribute.CodeTable = lnv_codetable
				
			case "checkbox"
				lnv_attribute = anv_dwattributefactory.CreateCheckboxAttribute()
				
				ls_tmp = ads_source.Describe(as_controlname + ".Values")
				
				this.ino_codetabletools.GetValueMap(&
						ls_tmp,&
						ref lnv_codetable,&
						ref ls_error)
						
				if not IsNull(ls_error) then
					as_error = "Could not parse codetable for object " + as_controlname + ". " + ls_error
					return lnv_attribute
				end if
						
				ls_codetabledisplay = this.ino_codetabletools.GetDisplayValues(lnv_codetable)
				ls_codetabledata = this.ino_codetabletools.GetDataValues(lnv_codetable)
				lnv_attribute.label = ls_codetabledisplay[1]
				lnv_attribute.CheckedValue = ls_codetabledata[1]
				lnv_attribute.UncheckedValue = ls_codetabledata[2]
				lnv_attribute.LeftText = ads_source.Describe(as_controlname + ".checkbox.lefttext") = "yes"
			case else
				lnv_attribute = anv_dwattributefactory.CreateTextAttributes()
		end choose


		
		ls_coltype = ads_source.Describe(as_controlname + ".ColType")
		of_GetControlFormat(ads_source, as_controlName, al_index, ls_format)
		
		if ls_coltype = "long" or ls_coltype = "number" or ls_coltype = "real" or ls_coltype = "ulong" then
			ls_unformattedValue = string(ads_source.GetItemNumber(al_index, as_controlname))
			if 	ls_editStyle = "dddw" &
				or ls_editStyle = "ddlb" then
				ls_outputText = ads_source.Describe("Evaluate('LookupDisplay(" + as_controlname + ")', " + string(al_index) + ")")
			else
				ls_outputText = string(ads_source.GetItemNumber(al_index, as_controlname), ls_format)
				if Left(ls_format, 1) = "$" then
					ls_datatypeflag = "money"
				else
					ls_datatypeflag = "number"
				end if
			end if

		elseif Match(ls_coltype, "char.*") then
			ls_unformattedValue = ads_source.GetItemString(al_index, as_controlname)
			if ls_format <> "[general]" then
				if Pos(ls_format, "#") <> 0 then
					ls_outputText = string(integer(ads_source.GetItemFormattedString(al_index, as_controlname)), ls_format)
				else
					ls_outputText = ads_source.GetItemFormattedString(al_index, as_controlname)
				end if
			else
				if 		ads_source.Describe(as_controlname + ".Edit.Style") = "dddw" &
						or ads_source.Describe(as_controlname + ".Edit.Style") = "ddlb" then
					ls_outputText = ads_source.Describe("Evaluate('LookupDisplay(" + as_controlname + ")', " + string(al_index) + ")")
				else
					ls_outputText = ads_source.GetItemString(al_index, as_controlname)
				end if

			end if

		elseif Match(ls_coltype, "decimal.*") then
			ls_unformattedValue = string(ads_source.GetItemDecimal(al_index, as_controlname))
			if Left(ls_format, 1) = "$" then
				ls_datatypeflag = "money"
			else
				ls_datatypeflag = "number"
			end if
			ls_outputText = string(ads_source.GetItemDecimal(al_index, as_controlname), ls_format)
		elseif ls_coltype = "datetime" then
			ls_datatypeflag = "datetime"
			ls_outputText = string(ads_source.GetItemDateTime(al_index, as_controlname), ls_format)
			ls_unformattedValue = string(ads_source.GetItemDateTime(al_index, as_controlname))
		elseif ls_coltype = "date" then
			ls_datatypeflag = "date"
			ls_outputText = string(ads_source.GetItemDate(al_index, as_controlname), ls_format)
			ls_unformattedValue = string(ads_source.GetItemDate(al_index, as_controlname))
		elseif ls_coltype = "time" then
			ls_datatypeflag = "time"
			ls_outputText = string(ads_source.GetItemTime(al_index, as_controlname), ls_format)
			ls_unformattedValue = string(ads_source.GetItemTime(al_index, as_controlname))
		end if
		
		if IsNull(ls_outputtext) then ls_outputtext = ""
text_common:

		lnv_attribute.Text = ls_outputText
		lnv_attribute.IsVisible = (of_GetValueForCurrentRow(ads_source, as_controlName, "Visible", al_index)) = "1"
		lnv_attribute.Alignment= integer(of_GetValueForCurrentRow(ads_source, as_controlName, "Alignment", al_index))
		lnv_attribute.FontSize = -byte(of_GetValueForCurrentRow(ads_source, as_controlName, "Font.Height", al_index))
		lnv_attribute.FontWeight = integer(of_GetValueForCurrentRow(ads_source, as_controlName, "Font.Weight", al_index))
		lnv_attribute.FontFace = of_GetValueForCurrentRow(ads_source, as_controlName, "Font.Face", al_index)
		lnv_attribute.Italics = (of_GetValueForCurrentRow(ads_source, as_controlName, "Font.Italic", al_index)) = "1"
		lnv_attribute.Underline = (of_GetValueForCurrentRow(ads_source, as_controlName, "Font.Underline", al_index)) = "1"
		lnv_attribute.Strikethrough= (of_GetValueForCurrentRow(ads_source, as_controlName, "Font.Strikethrough", al_index)) = "1"
		lnv_attribute.FormatString = ls_format
		lnv_attribute.RawText = ls_unformattedValue
		choose case ls_datatypeflag 
			case "money"
				lnv_attribute.DataType = 2
			case "number"
				lnv_attribute.DataType = 1
		end choose
		if istr_settings.b_exportcolor then 
			lnv_attribute.BackgroundColor = of_color2dwcolor(long(of_GetValueForCurrentRow(ads_source, as_controlName, "Background.Color", al_index)))
			lnv_attribute.FontColor = of_color2dwcolor(long(of_GetValueForCurrentRow(ads_source, as_controlName, "Color", al_index)))
		end if
	case "compute"
		goto case_column
		
	case "rectangle"
		lby_shapeType = byte(0)
case_shapecommon:
		lnv_attribute = anv_dwattributeFactory.CreateShapeAttributes()
		ls_tmp = of_GetValueForCurrentRow(ads_source, as_controlName, "Pen.Width", al_index)
		if IsNumber(ls_tmp) then
			lnv_attribute.OutlineWidth = UnitsToPixels(integer(ls_tmp), XUnitsToPixels!)
		end if
		
		lnv_attribute = anv_dwattributefactory.CreateShapeAttributes()
		lnv_attribute.IsVisible = (of_GetValueForCurrentRow(ads_source, as_controlName, "Visible", al_index)) = "1"
		lnv_attribute.Shape = lby_shapeType
		lnv_attribute.FillStyle = 		byte(of_GetValueForCurrentRow(ads_source, as_controlName, "Brush.Hatch", al_index))
		lnv_attribute.OutlineStyle = 	byte(of_GetValueForCurrentRow(ads_source, as_controlName, "Pen.Style", al_index))
		
		if istr_settings.b_exportcolor then
			lnv_attribute.OutlineColor = of_color2dwcolor(	long(of_GetValueForCurrentRow(ads_source, as_controlName, "Pen.Color", al_index)))
			lnv_attribute.FillColor = of_color2dwcolor(		long(of_GetValueForCurrentRow(ads_source, as_controlName, "Brush.Color", al_index)))
		else
			
			lnv_attribute.FillColor = of_color2dwcolor(16777215)
			lnv_attribute.OutlineColor = of_color2dwcolor(0)
		end if

	case "ellipse"
		lby_shapeType = byte(1)
		goto case_shapecommon
	case "line"
		lnv_attribute = anv_dwattributeFactory.CreateLineAttributes()
		lnv_attribute.IsVisible = (of_GetValueForCurrentRow(ads_source, as_controlName, "Visible", al_index)) = "1"

		lnv_attribute.LineStyle = byte(of_GetValueForCurrentRow(ads_source, as_controlName, "Pen.Style", al_index))
		lnv_attribute.SetStart(&
			UnitsToPixels(long(of_GetValueForCurrentRow(ads_source, as_controlName, "X1", al_index)), XUnitsToPixels!),&
			UnitsToPixels(long(of_GetValueForCurrentRow(ads_source, as_controlName, "Y1", al_index)), YUnitsToPixels!)&
		)
		lnv_attribute.SetEnd(&
			UnitsToPixels(long(of_GetValueForCurrentRow(ads_source, as_controlName, "X2", al_index)), XUnitsToPixels!),&
			UnitsToPixels(long(of_GetValueForCurrentRow(ads_source, as_controlName, "Y2", al_index)), YUnitsToPixels!)&
		)
		ls_tmp = of_GetValueForCurrentRow(ads_source, as_controlName, "Pen.Width", al_index)
		if IsNumber(ls_tmp) then
			lnv_attribute.LineWidth = UnitsToPixels(integer(ls_tmp), XUnitsToPixels!)
		end if
		
		if istr_settings.b_exportcolor then
			lnv_attribute.LineColor = of_color2dwcolor(long(of_GetValueForCurrentRow(ads_source, as_controlName, "Pen.Color", al_index)))
		else
			lnv_attribute.LineColor = of_color2dwcolor(0)
		end if

	case "bitmap"
		lnv_attribute = anv_dwAttributeFactory.CreatePictureAttributes()
		lnv_attribute.IsVisible = (of_GetValueForCurrentRow(ads_source, as_controlName, "Visible", al_index)) = "1"
		lnv_attribute.Filename = of_GetValueForCurrentRow(ads_source, as_controlName, "Filename", al_index)
		lnv_attribute.Transparency= byte(of_GetValueForCurrentRow(ads_source, as_controlName, "Transparency", al_index))
	case "button"	
		lnv_attribute = anv_dwAttributeFactory.CreateButtonAttributes()
		lnv_attribute.IsVisible = (of_GetValueForCurrentRow(ads_source, as_controlName, "Visible", al_index)) = "1"
		lnv_attribute.Text = of_GetValueForCurrentRow(ads_source, as_controlName, "Text", al_index)
		
end choose
endchoose:

return lnv_attribute
end function

private function string of_evaluateexpression (datastore ads_source, string as_expression, long al_row);return ads_source.Describe("Evaluate('" + as_expression + "', " + string(al_row) + ")")


end function

private function boolean of_getcontrolformat (datastore ads_source, string as_control, long al_row, ref string as_formatstring);// Returns whether the control has a format string
// and sends back such string on as_FormatString

string ls_format 

ls_format = of_GetValueForCurrentRow(ads_source, as_control, "EditMask.Mask", al_row)

if ls_format = "!" or ls_format = "?" or of_GetValueForCurrentRow(ads_source, as_control, "EditMask.UseFormat", al_row) = "yes" then
	ls_format = of_GetValueForCurrentRow(ads_source, as_control, "Format", al_row)
end if

if ls_format <> "" then
	as_formatstring = ls_format
	return true
end if

return false
end function

private function string of_getvalueforcurrentrow (datastore ads_source, string as_control, string as_property, long al_row);string ls_value
string ls_exp
if this.of_GetPropertyExpression(ads_source, as_control, as_property, ls_value, ls_exp) then
	// check if format string has expression
	ls_value = this.of_EvaluateExpression(ads_source, ls_exp, al_row)
end if

return ls_value
end function

public function boolean of_getpropertyexpression (datastore ads_source, string as_control, string as_property, ref string as_value, ref string as_expression);
string ls_propertyValue
string ls_aux[] // holds the result of the split funtion

ls_propertyValue = ads_source.Describe(as_control + "." + as_property)
if Left(ls_propertyValue, 1) = '"' and Right(ls_propertyValue, 1) = '"' then
	ls_propertyValue = Mid(ls_propertyValue, 2, Len(ls_propertyValue) - 2)
end if

this.ino_stringextensions.Split(ls_propertyValue, "~t", ref ls_aux)
as_value = ls_aux[1]
if UpperBound (ls_aux) > 1 then
	as_expression = ls_aux[2]
	return true
end if

return false

end function

public function integer of_convert (readonly datastore ads_ds, readonly string as_path);///// Exports a DataWindow as an XLSX file
/////
/////	- ads_ds: the datawindow
/////	- as_path: the path to which to export the file
/////
///// - returns dotnetobject (List<ExportedCellBase>) containing all the exported cells

string ls_error
powerobject ls_callback
integer li_return

SetNull(ls_error)
SetNull(ls_callback)
li_return = of_convert(ads_ds, as_path, false, "", ref ls_error, ls_callback, "")

if not IsNull(ls_error) then
	SetNull(li_return)
	MessageBox("Could not convert DataWindow", ls_error)
	return -1
end if

return 1
end function

public function integer of_convert (datastore ads_ds, string as_path, boolean ab_append, string as_sheetname, ref string as_error, powerobject ao_callback, string as_callbackevent);///// Exports a DataStore as an XLSX file
/////
/////	- ads_ds: the datastore
/////	- as_path: the path to which to export the file
/////	- ab_append: whether to append the sheet to an existing file
/////	- as_sheetname: the new sheet's name
/////	- as_error: string containing the error message if it fails
/////	- ao_callback: callback object to which the progress will be reported
/////	- as_callbackevent: the event that will be invoked from ao_callback
/////
///// - returns dotnetobject (List<ExportedCellBase>) containing all the exported cells

dotnetobject 								ldn_virtualGrid
dotnetobject 								matrix
dotnetobject 								lno_XlsxWriter
dotnetobject 								lno_aggregatedControls
dotnetobject 								lno_aggregatedControlsCurrent
n_cst_controlmatrixbuilder 				lno_builder
n_cst_virtualgridbuilder 					lno_gridbuilder
n_cst_dwattributefactory 				lno_attribFactory
n_cst_datasetbuilder 						lno_DataSetBuilder
n_cst_VirtualGridXlsxWriterBuilder 	lno_writerBuilder
n_cst_aggregators 						lno_aggregators
int 											li_error
int												li_ret
long 											ll_rowCount 
long 											i,j
boolean 										lb_bandexcluded
string 										ls_error
string 										ls_controls[]
string 										ls_excludedbands[]
string											ls_unsupportedObjects[]
string 										ls_aux

SetNull(lno_aggregatedControls)

if Trim(as_path) = "" then
	as_error = "No path specified"
	
	ino_conversionResult = lno_aggregatedControls
	return -1
end if

if ab_append and Trim(as_sheetname) = "" then
	as_error = "Did not specify a sheet name"
	ino_conversionResult = lno_aggregatedControls
	return -1
end if

try
	if istr_settings.b_logging then
		ino_logger = create n_cst_logger
	else
		ino_logger = create n_cst_nulllogger
	end if
	li_ret = ino_logger.of_init("logs\" &
		+ ino_stringextensions.Replace(string(Today(), "yyyy-mm-dd"), "-", "") &
		+ ino_stringextensions.Replace(string(Now(), "hh:mm:ss"), ":", "") + ".log")
		
	if(li_ret = -1 ) then
		MessageBox("Logger error", "Could not initialize logger")
		as_error = "Could not initialize logger"
		ino_conversionResult = lno_aggregatedControls
		return -1
	end if
	
	ino_logger.of_log("Begin conversion")
	li_error = of_Init(ls_error)
	if li_error <> 0 then
		as_error = "Execution Context Init: " + ls_error
		ino_logger.of_error(as_error)
		ino_conversionResult = lno_aggregatedControls
		return -1
	end if
	
	
	SetNull(lno_aggregators)
	SetNull(lno_aggregatedControls)
	SetNull(lno_aggregatedControlsCurrent)
	lno_aggregators = create n_cst_aggregators
	lno_aggregators.of_createondemand( )
	if lno_aggregators.il_errornumber <> 0 then
		ino_logger.of_error(lno_aggregators.is_errortext)
		as_error = "Could not create aggregator: " + lno_aggregators.is_errortext
		ino_logger.of_error(ls_error)
		ino_conversionResult = lno_aggregatedControls
		return -1
	end if
	
	lno_builder = create n_cst_controlmatrixbuilder
	
	SetNull(ls_error)
	
	for i = 1 to UpperBound(is_bandstoinclude) 
		lb_bandexcluded = true
		for j = 1 to UpperBound(istr_settings.s_bands)
			if istr_settings.s_bands[j] = is_bandstoinclude[i] then
				lb_bandexcluded = false
			end if
		next
		
		if lb_bandexcluded then ls_excludedbands[UpperBound(ls_excludedbands) + 1] = is_bandstoinclude[i]
	next

	ino_logger.of_log("Building control matrix")
	lno_builder.of_buildcontrolmatrix(this, ads_ds, integer(ads_ds.Object.DataWindow.Units), ref matrix, ref ls_error, ref ls_controls, ref ls_unsupportedObjects, ls_excludedbands )
	
	if UpperBound(ls_unsupportedObjects) > 0 then
		ls_aux = ino_ObjectExtensions.ToString(ls_unsupportedObjects)
		for i = 1 to UpperBound(ls_unsupportedobjects)
			ino_logger.of_warning("Unsupported object: " + ls_unsupportedObjects[i])
		next
	end if

	if not IsNull(ls_error) then
		MessageBox("Error", ls_error)
		ino_logger.of_error(ls_error)
		ino_conversionResult = lno_aggregatedControls
		return -1
	end if
	
	ino_logger.of_log("Creating Virtual Grid Builder")
	lno_gridbuilder = create n_cst_virtualgridbuilder 
	lno_gridbuilder.of_Createondemand( )
	if lno_gridbuilder.il_errornumber <> 0 then
		ino_logger.of_error(lno_gridbuilder.is_errortext)
		as_error = "Could not create grid builder: " + lno_gridbuilder.is_errortext
		ino_logger.of_error(ls_error)
		ino_conversionResult = lno_aggregatedControls
		return -1
	end if
	
	
	/// Threshold values determine how close controls can be to each other
	/// before they're considered to belong to the same row/column
	lno_gridbuilder.XThreshold = istr_settings.i_xthreshold
	lno_gridbuilder.YThreshold = istr_settings.i_ythreshold
	

	setNull(ldn_virtualGrid)
	
	try 
		ino_logger.of_log("Building VirtualGrid")
		ldn_virtualGrid = lno_gridBuilder.Build(&
				matrix,&
				byte(ads_ds.Object.DataWindow.Processing),&
				true,&
				istr_settings.b_ignoremisaligned,&
				ref ls_error)
	catch (RuntimeError e)
		ino_logger.of_error(e.GetMessage())
		ls_error = e.GetMessage()
	end try
	if not IsNull(ls_error) then
		as_error = "Error creating VirtualGrid: " + ls_error
		ino_logger.of_error(as_error)
		ino_conversionResult = lno_aggregatedControls
		return -1
	end if

	ino_logger.of_log("Creating DwAttributeFactory")
	lno_attribFactory = create n_cst_dwattributefactory
	lno_attribFactory.of_createondemand( )
	if lno_attribFactory.il_errornumber <> 0 then
			as_error = "AttributeFactory: " + lno_attribFactory.is_errortext
			ino_logger.of_error(as_error)
			ino_conversionResult = lno_aggregatedControls
			return -1
	end if

	ino_logger.of_log("Creating DataSetBuilder")
	lno_DataSetBuilder = create n_cst_datasetbuilder
	lno_DataSetBuilder.of_createondemand( )
	if lno_DataSetBuilder.il_errornumber <> 0 then
		as_error = "DataSetBuilder: " + lno_DataSetBuilder.is_errortext
		ino_logger.of_error(as_error)
		ino_conversionResult = lno_aggregatedControls
		return -1
	end if
	
	ino_logger.of_log("Creating VirtualGridXlsxWriterBuilder")
	lno_writerBuilder = create n_cst_VirtualGridXlsxWriterBuilder
	lno_writerBuilder.of_createondemand( )
	if lno_writerBuilder.il_errornumber <> 0 then
		as_error = "VirtualGridXlsxWriterBuilder: " + lno_writerBuilder.is_errortext
		ino_logger.of_error(as_error)
		ino_conversionResult = lno_aggregatedControls
		return -1
	end if
	
	SetNull(ls_error)
	
	lno_writerBuilder.WriteToPath = as_path
	lno_writerBuilder.Append = ab_append
	lno_writerBuilder.AppendToSheetName = as_sheetname


	ino_logger.of_log("Building VirtualGridXlsxWriter")
	lno_XlsxWriter = lno_writerBuilder.of_build( ldn_virtualGrid, ref ls_error)
	if not IsNull(ls_error) then
		as_error = "VirtualGridXlsxWriterBuilder.Build: " + ls_error
		ino_logger.of_error(as_error)
		ino_conversionResult = lno_aggregatedControls
		return -1
	end if

	ll_rowCount = ads_ds.Rowcount( )
	SetNull(ls_error)

	ino_logger.of_log("Beginning row processing")
	for i = 1 to ll_rowCount
		for j = 1 to UpperBound(ls_controls)
			lno_DataSetBuilder.of_addcontrolattribute(ls_controls[j], of_dwControlToAttribute(ads_ds, i, ls_controls[j], lno_attribFactory, ls_error))
			
			if not IsNull(ls_error) then
				as_error = "Error processing object attributes: " + ls_error
				ino_logger.of_error(as_error)
				ino_conversionResult = lno_aggregatedControls
				return -1
			end if
		next
		
		lno_aggregatedControlsCurrent = lno_XlsxWriter.EnterData(lno_dataSetBuilder.of_build(), ref ls_error)	
		if IsNull(lno_aggregatedControls) then
			lno_aggregatedControls = lno_aggregatedControlsCurrent
		else
			if lno_aggregatedControlsCurrent.Count > 0 then
				lno_aggregators.of_Merge( ref lno_aggregatedControls, lno_aggregatedControlsCurrent)
			end if
		end if
		if not IsNull(ls_error) then 
			as_error = ls_error
			ino_conversionResult = lno_aggregatedControls
			return -1
		end if
		lno_dataSetBuilder.of_clear()
		if IsValid(ao_callback) and not IsNull(ao_callback) then &
			ao_callback.TriggerEvent(as_callbackevent, i, ll_rowCount)
	next
	
	ino_logger.of_log("Finished processing rows")

	
	// Indicate that there's no more data, let the trailers be written
	SetNull(lno_DataSetBuilder)
	lno_aggregatedControlsCurrent = lno_XlsxWriter.EnterData(lno_dataSetBuilder, ref ls_error)	
	if IsNull(lno_aggregatedControls) then
		lno_aggregatedControls = lno_aggregatedControlsCurrent
	else
			if lno_aggregatedControlsCurrent.Count > 0 then
				lno_aggregators.of_Merge( ref lno_aggregatedControls, lno_aggregatedControlsCurrent)
			end if
	end if
	
	ino_logger.of_log("Writing trailers")
	
	if not IsNull(ls_error) then
		as_error = "Could not write trailers: " + ls_error
		ino_logger.of_error(as_error)
	end if

	ino_logger.of_log("Writing file")
	if not lno_XlsxWriter.Write(	as_path, ref ls_error) then
		as_error = "Could not write file: " +  ls_error
		ino_logger.of_error(as_error)
	end if
	  	
	ino_conversionResult = lno_aggregatedControls
	return 1
catch (Throwable ex)
	as_error = "Exception caught: " + ex.GetMessage()
finally
	if ls_error <> "" then
		ino_logger.of_error("Unhandled error: " + ls_error)
	end if

	if not IsNUll(lno_XlsxWriter) and IsValid(lno_XlsxWriter) then
		lno_XlsxWriter.Dispose()
	end if

	destroy lno_XlsxWriter
	destroy ino_logger
	destroy lno_builder
	destroy lno_gridbuilder
	destroy lno_attribFactory
	destroy lno_DataSetBuilder
	destroy lno_writerBuilder
end try



return -1
end function

public function integer of_convert (readonly datawindow adw_dw, readonly string as_path);///// Exports a DataWindow as an XLSX file
/////
/////	- adw_dw: the datawindow
/////	- as_path: the path to which to export the file
/////
///// - returns dotnetobject (List<ExportedCellBase>) containing all the exported cells

datastore lds_ds

lds_ds = f_dw2ds(adw_dw, Primary!)
return of_convert(lds_ds, as_path)
end function

public function integer of_convert (readonly datawindow adw_dw, readonly string as_path, boolean ab_append, readonly string as_sheetname, ref string as_error, readonly powerobject ao_callback, readonly string as_callbackevent);///// Exports a DataStore as an XLSX file
/////
/////	- adw_dw: the datastore
/////	- as_path: the path to which to export the file
/////	- ab_append: whether to append the sheet to an existing file
/////	- as_sheetname: the new sheet's name
/////	- as_error: string containing the error message if it fails
/////	- ao_callback: callback object to which the progress will be reported
/////	- as_callbackevent: the event that will be invoked from ao_callback
/////
///// - returns dotnetobject (List<ExportedCellBase>) containing all the exported cells

datastore lds_ds

lds_ds = f_dw2ds(adw_dw, Primary!)

return of_convert(lds_ds, as_path, ab_append, as_sheetname, as_error, ao_callback, as_callbackevent)
end function

on n_cst_dw2excelconverter.create
call super::create
TriggerEvent( this, "constructor" )
end on

on n_cst_dw2excelconverter.destroy
TriggerEvent( this, "destructor" )
call super::destroy
end on

