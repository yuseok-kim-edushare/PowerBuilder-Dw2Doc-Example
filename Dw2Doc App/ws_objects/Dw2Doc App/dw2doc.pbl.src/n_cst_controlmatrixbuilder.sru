$PBExportHeader$n_cst_controlmatrixbuilder.sru
forward
global type n_cst_controlmatrixbuilder from nonvisualobject
end type
end forward

global type n_cst_controlmatrixbuilder from nonvisualobject
end type
global n_cst_controlmatrixbuilder n_cst_controlmatrixbuilder

forward prototypes
public function string of_getdwcontrolprop (readonly datastore ads_ds, string as_control, string as_property)
public subroutine of_getdatawindowobjectnames (n_cst_dw2excelconverter ano_converter, readonly datastore ads_source, ref string as_objects[])
public function long of_getcontrolbandheight (datastore ads_datastore, string as_control)
public function integer of_buildcontrolmatrix (n_cst_dw2excelconverter ano_ctx, readonly datastore ads_source, integer ai_units, ref dotnetobject anv_controlmatrix, ref string as_error, ref string as_controls[], ref string as_unsupportedcontrols[], readonly string as_excludedbands[])
end prototypes

public function string of_getdwcontrolprop (readonly datastore ads_ds, string as_control, string as_property);return ads_ds.Describe(as_control + "." + as_property) 
end function

public subroutine of_getdatawindowobjectnames (n_cst_dw2excelconverter ano_converter, readonly datastore ads_source, ref string as_objects[]);string ls_objects

ls_objects = ads_source .Describe("DataWindow.VisualObjects")

ano_converter.ino_StringExtensions.Split(ls_objects, "~t", ref as_objects) 
end subroutine

public function long of_getcontrolbandheight (datastore ads_datastore, string as_control);return long(ads_datastore.Describe("DataWindow." + of_getDwControlProp( ads_datastore, as_control, "band")  + ".Height"))
end function

public function integer of_buildcontrolmatrix (n_cst_dw2excelconverter ano_ctx, readonly datastore ads_source, integer ai_units, ref dotnetobject anv_controlmatrix, ref string as_error, ref string as_controls[], ref string as_unsupportedcontrols[], readonly string as_excludedbands[]);n_cst_dwcontrolmatrixbuilder lnv_controlMatrixBuilder
string ls_band
string ls_dwobjects[]
string ls_dwfilteredobjects[]
string ls_unhandledTypes[]
string ls_type
string ls_converted 
string ls_error
string ls_controlVisible
string ls_controlVisibleExpression
boolean lb_floating
double ld_yp
int i
int j

dotnetobject lnv_controlmatrix
dotnetobject lnv_bands

lnv_controlMatrixBuilder = create n_cst_dwcontrolmatrixbuilder
lnv_controlMatrixBuilder.of_CreateOnDemand()

if lnv_controlMatrixBuilder.il_errornumber <> 0 then
	as_error=  "Could not create MatrixBuilder: " + lnv_controlMatrixBuilder.is_errortext
	return -1
end if

if ai_units = 0 then
	ld_yp = PixelsToUnits(1, YPixelsToUnits!)
elseif ai_units = 1 then
	ld_yp = 1.0
else
	as_error = "Could not create MatrixBuilder: Unsupported units (" + string(ai_units) + ")"
	return -1 
end if

lnv_bands = ano_ctx.ino_DwTools.GetBands(ads_source.Object.DataWindow.Syntax, ld_yp, as_excludedbands)

of_getdatawindowobjectnames(ano_ctx, ads_source, ref ls_dwobjects)

lb_floating = false

for i = 1 to UpperBound(ls_dwobjects)
	ls_band = of_getDwControlProp(ads_source, ls_dwobjects[i], "band")
	
	for j = 1 to UpperBound(as_excludedbands)
		/// Skip controls in excluded bands
		if ls_band = as_excludedbands[j] then goto end_object_loop
	next
	
	// check if control is statically invisible (visible = 0 and no expression)
	ano_ctx.of_getPropertyExpression(ads_source, ls_dwobjects[i], "Visible", ref ls_controlVisible, ref ls_controlVisibleExpression)
	if ls_controlVisible = "0" and ls_controlVisibleExpression = '' then
		// don't include it in the grid calculations
		continue
	end if
	
	ls_type = ads_source.Describe(ls_dwobjects[i] + ".Type")
	choose case ls_type 
		case "column"
			lb_floating = ads_source.Describe(ls_dwobjects[i] + ".Tag") = "dw2xlsx_floating"
object_2d:
			if ads_source.Describe(ls_dwobjects[i] + ".BitmapName") = 'yes' and &
				not ano_ctx.istr_settings.b_renderbitmap then 
				goto end_object_loop
			end if
			ls_dwfilteredobjects[UpperBound(ls_dwfilteredobjects) + 1] = ls_dwobjects[i]
			if ai_units = 1 then
				lnv_controlMatrixBuilder.Add2dcontrol(&
					ls_dwobjects[i], & 
					ls_band,&
					long(of_getDwControlProp(ads_source, ls_dwobjects[i], "X")),&
					long(of_getDwControlProp(ads_source, ls_dwobjects[i], "Y")),&
					long(of_getDwControlProp(ads_source, ls_dwobjects[i], "Width")),&
					long(of_getDwControlProp(ads_source, ls_dwobjects[i], "Height")),&
					lb_floating)	
			else
				lnv_controlMatrixBuilder.Add2dcontrol(&
					ls_dwobjects[i], & 
					ls_band,&
					UnitsToPixels(long(of_getDwControlProp(ads_source, ls_dwobjects[i], "X")), XUnitsToPixels!),&
					UnitsToPixels(long(of_getDwControlProp(ads_source, ls_dwobjects[i], "Y")), YUnitsToPixels!),&
					UnitsToPixels(long(of_getDwControlProp(ads_source, ls_dwobjects[i], "Width")),XUnitsToPixels!),&
					UnitsToPixels(long(of_getDwControlProp(ads_source, ls_dwobjects[i], "Height")), YUnitsToPixels!),&
					lb_floating)	
			end if

		case "text"
			lb_floating = ads_source.Describe(ls_dwobjects[i] + ".Tag") = "dw2xlsx_floating"
			goto object_2d
		case "compute"
			lb_floating = ads_source.Describe(ls_dwobjects[i] + ".Tag") = "dw2xlsx_floating"
			goto object_2d
		case "rectangle"
			lb_floating = true
			if ano_ctx.istr_settings.b_renderrectangle then
				goto object_2d
			end if
		case "ellipse"
			lb_floating = true
			if ano_ctx.istr_settings.b_renderoval then
				goto object_2d
			end if
		case "line"
			if ano_ctx.istr_settings.b_renderline then
				ls_dwfilteredobjects[UpperBound(ls_dwfilteredobjects) + 1] = ls_dwobjects[i]
				if ai_units = 1 then
					lnv_controlMatrixBuilder.Add1DControl(ls_dwobjects[i], &
						ls_band,&
							long(of_getDwControlProp(ads_source, ls_dwobjects[i], "X1")),&
							long(of_getDwControlProp(ads_source, ls_dwobjects[i], "Y1")),&
							long(of_getDwControlProp(ads_source, ls_dwobjects[i], "X2")) ,&
							long(of_getDwControlProp(ads_source, ls_dwobjects[i], "Y2")),&
							true) // lb_floating	
				else
					lnv_controlMatrixBuilder.Add1DControl(ls_dwobjects[i], &
						ls_band,&
							UnitsToPixels(long(of_getDwControlProp(ads_source, ls_dwobjects[i], "X1")), XUnitsToPixels!),&
							UnitsToPixels(long(of_getDwControlProp(ads_source, ls_dwobjects[i], "Y1")), YUnitsToPixels!),&
							UnitsToPixels(long(of_getDwControlProp(ads_source, ls_dwobjects[i], "X2")), XUnitsToPixels!) ,&
							UnitsToPixels(long(of_getDwControlProp(ads_source, ls_dwobjects[i], "Y2")), YUnitsToPixels!),&
							true) // lb_floating	
				end if

			end if
			
		case "bitmap"
			lb_floating = true
			if ano_ctx.istr_settings.b_renderbitmap then
				goto object_2d
			end if
		case "button"
			lb_floating = true
			if ano_ctx.istr_settings.b_renderbutton then
				goto object_2d
			end if
//		case "groupbox"
//			// groupboxes are commonly used but there's no standard support for them on NPOI
//			// so we just accept them but do nothing with them. Maybe they can be implemented 
//			// as a composite shape? (Square + label)
		case else 
			ls_unhandledTypes[UpperBound(ls_unhandledTypes) + 1] = ls_dwobjects[i] + ":" + ls_type
	end choose
end_object_loop:
next

as_controls = ls_dwfilteredobjects
ls_dwobjects = ls_dwfilteredobjects

//ls_converted = ano_ctx.ino_ObjectExtensions.ToString(ls_unhandledTypes)
//if UpperBound(ls_unhandledTypes) > 0 then
//	MessageBox("Unhandled objects", ls_converted)
//end if

as_unsupportedControls = ls_unhandledTypes

lnv_controlMatrixBuilder.of_setBands(lnv_bands)

lnv_controlmatrix = lnv_controlMatrixBuilder.Build(ref ls_error)

if not IsNull(ls_error) then
	as_error = ls_error
	return -3
end if

anv_controlmatrix = lnv_controlmatrix
return 1
end function

on n_cst_controlmatrixbuilder.create
call super::create
TriggerEvent( this, "constructor" )
end on

on n_cst_controlmatrixbuilder.destroy
TriggerEvent( this, "destructor" )
call super::destroy
end on

