using Appeon.DotnetDemo.Dw2Doc.Common.DwObjects.DwObjectAttributes;
using Appeon.DotnetDemo.Dw2Doc.Common.Extensions;
using Appeon.DotnetDemo.Dw2Doc.Common.VirtualGrid;
using Appeon.DotnetDemo.Dw2Doc.Common.VirtualGridWriter.Abstractions;
using Appeon.DotnetDemo.Dw2Doc.Common.VirtualGridWriter.Models;
using Appeon.DotnetDemo.Dw2Doc.Xlsx.Extensions;
using Appeon.DotnetDemo.Dw2Doc.Xlsx.Helpers;
using Appeon.DotnetDemo.Dw2Doc.Xlsx.VirtualGridWriter.Renderers;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using System.Diagnostics.CodeAnalysis;
using System.Text.RegularExpressions;

namespace Appeon.DotnetDemo.Dw2Doc.Xlsx.VirtualGridWriter.XlsxWriter
{
    public class VirtualGridXlsxWriter : AbstractVirtualGridWriter
    {
        private readonly Regex mergedCellsRegex = new(@"Cannot add merged region (.+) to sheet because it overlaps with an existing merged region \((.+)\)\.");
        private int _startRowOffset = 0;
        private int _startColumnOffset = 0;
        private int _currentRowOffset;
        private XSSFWorkbook? _workbook;
        private XSSFSheet? _sheet;
        private XSSFDrawing? _drawingPatriarch;
        private ISet<ColumnDefinition>? _resizedColumns;
        private IDictionary<string, int>? _pictureCache;
        private bool _writerInitialized = false;
        private string _path;
        private FileStream? _destinationFile;
        private bool _closed = false;
        private Dictionary<string, VirtualCell?> rangedCells = new();

        internal VirtualGridXlsxWriter(VirtualGrid grid, string path)
            : base(grid)
        {
            _workbook = new XSSFWorkbook();
            _sheet = (XSSFSheet)_workbook.CreateSheet();
            Console.WriteLine($"Created new sheet with default name: {_workbook.GetSheetName(0)}");
            
            _drawingPatriarch = (XSSFDrawing)_sheet.CreateDrawingPatriarch();
            _resizedColumns = new HashSet<ColumnDefinition>();

            _resizedColumns.AddRange(VirtualGrid.Columns);
            _pictureCache = new Dictionary<string, int>();
            _path = path;
            LockFile(_path);
            
            Console.WriteLine($"Initialized VirtualGridXlsxWriter with path: {path}");
        }

        internal VirtualGridXlsxWriter(VirtualGrid grid, string path, string sheetname) : base(grid)
        {
            _path = path;
            Console.WriteLine($"Opening existing file: {path}");
            
            try 
            {
                if (!File.Exists(path))
                {
                    throw new FileNotFoundException($"File not found: {path}");
                }
                
                var fileStream = new FileStream(path, FileMode.Open, FileAccess.Read);
                _workbook = new XSSFWorkbook(fileStream);
                fileStream.Close();
                
                Console.WriteLine($"Opened workbook with {_workbook.NumberOfSheets} sheets");
                
                if (_workbook.GetSheet(sheetname) is not null)
                {
                    throw new ArgumentException($"Sheet '{sheetname}' already exists", nameof(sheetname));
                }

                _sheet = (XSSFSheet)_workbook.CreateSheet(sheetname);
                _workbook.SetActiveSheet(_workbook.GetSheetIndex(_sheet));
                Console.WriteLine($"Created new sheet: {sheetname}");

                _drawingPatriarch = (XSSFDrawing)_sheet.CreateDrawingPatriarch();
                _resizedColumns = new HashSet<ColumnDefinition>();

                _resizedColumns.AddRange(VirtualGrid.Columns);
                _pictureCache = new Dictionary<string, int>();
                
                LockFile(_path);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error initializing writer: {ex.Message}");
                throw;
            }
        }

        [MemberNotNull(nameof(_destinationFile))]
        private void LockFile(string filePath)
        {
            try
            {
                if (File.Exists(filePath))
                {
                    Console.WriteLine($"Opening existing file for writing: {filePath}");
                    // Use FileMode.Create to truncate and create a new file, even if it exists
                    _destinationFile = new FileStream(filePath, FileMode.Create, FileAccess.Write, FileShare.None);
                }
                else
                {
                    Console.WriteLine($"Creating new file: {filePath}");
                    _destinationFile = new FileStream(filePath, FileMode.Create, FileAccess.Write, FileShare.None);
                }
                
                Console.WriteLine($"File locked for writing with {(File.Exists(filePath) ? "FileMode.Create" : "FileMode.Create")}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error locking file: {ex.Message}");
                throw;
            }
        }

        private void InitWriter()
        {
            if (_writerInitialized) return;
            
            if (_sheet != null)
            {
                var row = _sheet.CreateRow(_startRowOffset);
                foreach (var column in _resizedColumns ?? Enumerable.Empty<ColumnDefinition>())
                {
                    row.CreateCell(column.IndexOffset + _startColumnOffset);
                    _sheet.SetColumnWidth(column.IndexOffset + _startColumnOffset, (int)UnitConversion.PixelsToColumnWidth(column.Size));
                }

                _sheet.RemoveRow(row);
            }

            _writerInitialized = true;
        }

        protected override IList<ExportedCellBase>? WriteRows(IList<RowDefinition> rows, IDictionary<string, DwObjectAttributesBase>? data)
        {
            InitWriter();
            if (data is null)
            {
                Console.WriteLine("Data is null, cannot write rows");
                return null;
            }

            try
            {
                if (_closed)
                {
                    throw new InvalidOperationException("This writer has already been closed");
                }

                if (_sheet == null || _workbook == null)
                {
                    throw new InvalidOperationException("Sheet or workbook is null");
                }

                var exportedCells = new List<ExportedCellBase>();
                ExportedCellBase? exportedCell = null;
                
                Console.WriteLine($"Processing {rows.Count} rows");
                
                foreach (var row in rows)
                {
                    if (_sheet == null) continue;
                    
                    // Create a physical row in the sheet
                    var xRow = _sheet.CreateRow(_startRowOffset + _currentRowOffset);
                    Console.WriteLine($"Created row at index {_startRowOffset + _currentRowOffset}");

                    checked
                    {
                        xRow.Height = (short)(row.Size).PixelsToTwips(Windows.ScreenTools.MeasureDirection.Vertical);
                    }

                    ICell? previousCell = null;
                    int lastOccupiedColumn = 0;
                    int cellsProcessed = 0;
                    
                    // Combine regular and floating objects for processing
                    var allCells = row.Objects.Concat(row.FloatingObjects).ToList();
                    Console.WriteLine($"Processing {allCells.Count} cells in row");
                    
                    foreach (var cell in allCells)
                    {
                        if (!data.TryGetValue(cell.Object.Name, out var attribute))
                        {
                            Console.WriteLine($"Warning: No attribute found for cell {cell.Object.Name}");
                            continue;
                        }

                        if (!attribute.IsVisible)
                        {
                            Console.WriteLine($"Cell {cell.Object.Name} is not visible, skipping");
                            continue;
                        }
                        
                        var renderer = RendererLocator.Find(attribute.GetType());
                        if (renderer == null)
                        {
                            Console.WriteLine($"Warning: Could not find renderer for attribute type {attribute.GetType().FullName}");
                            continue;
                        }

                        if (renderer is XlsxPictureRenderer pictureRenderer && _pictureCache != null)
                        {
                            pictureRenderer.SetPictureCache(_pictureCache);
                        }

                        if (_sheet == null)
                        {
                            continue;
                        }

                        if (cell.OwningColumn is not null && !cell.Object.Floating)
                        { // cell is not floating
                            var xCell = xRow.CreateCell(cell.OwningColumn.IndexOffset + _startColumnOffset);
                            Console.WriteLine($"Created cell at column {cell.OwningColumn.IndexOffset + _startColumnOffset} for {cell.Object.Name}");

                            var style = _workbook.CreateCellStyle();
                            switch (attribute)
                            {
                                case DwTextAttributes txt:
                                    //style.Alignment = txt.Alignment.ToNpoiHorizontalAlignment();
                                    //xCell.CellStyle = style;
                                    break;
                            }

                            if (xCell.ColumnIndex - lastOccupiedColumn > 2)
                            {
                                var cellRange = new CellRangeAddress(
                                    xCell.RowIndex,
                                    xCell.RowIndex,
                                    lastOccupiedColumn + 1,
                                    xCell.ColumnIndex - 1);

                                rangedCells[cellRange.FormatAsString()] = cell;
                                _sheet.AddMergedRegion(cellRange);
                            }

                            lastOccupiedColumn = cell.OwningColumn.IndexOffset + _startColumnOffset;

                            previousCell = xCell;
                            exportedCell = (renderer.Render(_sheet, cell, attribute, xCell));

                            if (exportedCell is not null)
                            {
                                exportedCells.Add(exportedCell);
                                cellsProcessed++;
                                Console.WriteLine($"Successfully rendered cell {cell.Object.Name}");
                            }
                            else
                            {
                                Console.WriteLine($"Warning: Renderer returned null for cell {cell.Object.Name}");
                            }

                            lastOccupiedColumn = cell.OwningColumn.IndexOffset + _startColumnOffset;
                            if (cell.ColumnSpan > 1)
                            {
                                var cellRange = new CellRangeAddress(
                                    xCell.RowIndex,
                                    xCell.RowIndex,
                                    xCell.ColumnIndex,
                                    xCell.ColumnIndex + cell.ColumnSpan - 1);

                                rangedCells[cellRange.FormatAsString()] = cell;

                                lastOccupiedColumn += (cell.ColumnSpan - 1);

                                _sheet.AddMergedRegion(cellRange);
                            }
                        }
                        else
                        { // cell is floating
                            if (cell is not FloatingVirtualCell floatingCell)
                            {
                                Console.WriteLine("Warning: Non-floating cell in FloatingObjects list");
                                throw new InvalidOperationException("Non-floating cell in FloatingObjects list");
                            }
                            previousCell = null;
                            if (_drawingPatriarch != null)
                            {
                                exportedCell = renderer.Render(_sheet, floatingCell,
                                    data[cell.Object.Name],
                                    (_startColumnOffset + floatingCell.Offset.StartColumn.IndexOffset,
                                        _currentRowOffset,
                                        _drawingPatriarch));

                                if (exportedCell is not null)
                                {
                                    exportedCells.Add(exportedCell);
                                    cellsProcessed++;
                                    Console.WriteLine($"Successfully rendered floating cell {cell.Object.Name}");
                                }
                                else
                                {
                                    Console.WriteLine($"Warning: Renderer returned null for floating cell {cell.Object.Name}");
                                }
                            }
                        }
                    }

                    Console.WriteLine($"Processed {cellsProcessed} cells out of {allCells.Count} for row");

                    // Handle empty rows or filler rows
                    int unoccupiedTrailingColumns = 0;
                    if (row.IsFiller
                        && VirtualGrid.Columns.Count > 1
                        && row.Objects.Count == 0
                        && row.FloatingObjects.Count == 0
                        || (unoccupiedTrailingColumns = VirtualGrid.Columns.Count - lastOccupiedColumn) > 2)
                    {
                        if (_sheet != null)
                        {
                            var cellRange = new CellRangeAddress(
                                _currentRowOffset + _startRowOffset,
                                _currentRowOffset + _startRowOffset,
                                _startColumnOffset + unoccupiedTrailingColumns > 2 ? (lastOccupiedColumn + 1) : 0,
                                _startColumnOffset + VirtualGrid.Columns.Count - 1);
                            rangedCells[cellRange.FormatAsString()] = null;
                            _sheet.AddMergedRegion(cellRange);
                        }
                    }
                    
                    // If row has no cells, create at least one empty cell to ensure the row exists
                    if (cellsProcessed == 0 && xRow.PhysicalNumberOfCells == 0)
                    {
                        xRow.CreateCell(0);
                        Console.WriteLine("Created empty cell to ensure row exists");
                    }

                    ++_currentRowOffset;
                }
                
                Console.WriteLine($"Finished processing rows, created {exportedCells.Count} cells");
                return exportedCells;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Exception in WriteRows: {ex.Message}");
                _destinationFile?.Close();

                var match = mergedCellsRegex.Match(ex.Message);
                if (match.Success)
                {
                    throw new Exception($"Control {rangedCells[match.Groups[1].Value]?.Object.Name ?? "NULL"} overlaps " +
                        $"with object {rangedCells[match.Groups[2].Value]?.Object.Name ?? "NULL"}");
                }
                throw;
            }
        }

        public override bool Write(string? sheetname, out string? error)
        {
            error = null;
            
            if (_workbook == null)
            {
                error = "Workbook is null";
                return false;
            }
            if (_closed)
            {
                error = "This writer has already been closed";
                return false;
            }
            
            try
            {
                // Make sure we have a sheet to work with
                if (_sheet == null)
                {
                    _sheet = (XSSFSheet)_workbook.CreateSheet(sheetname ?? "Sheet1");
                    // Set sheet as active
                    _workbook.SetActiveSheet(_workbook.GetSheetIndex(_sheet));
                    Console.WriteLine($"Created sheet with name: {sheetname ?? "Sheet1"}");
                }
                else if (!string.IsNullOrEmpty(sheetname) && _workbook.GetSheetName(_workbook.GetSheetIndex(_sheet)) != sheetname)
                {
                    // Rename the sheet if needed
                    int sheetIndex = _workbook.GetSheetIndex(_sheet);
                    _workbook.SetSheetName(sheetIndex, sheetname);
                    Console.WriteLine($"Renamed sheet to: {sheetname}");
                }
                
                // Only process control attributes if sheet is empty and we weren't called from WriteEntireGrid
                // The StackTrace check helps us determine if we were called by WriteEntireGrid
                bool calledFromWriteEntireGrid = new System.Diagnostics.StackTrace().ToString().Contains("WriteEntireGrid");
                
                if (_sheet.PhysicalNumberOfRows == 0 && !calledFromWriteEntireGrid)
                {
                    var controlAttributesField = typeof(VirtualGrid).GetField("_controlAttributes", 
                        System.Reflection.BindingFlags.NonPublic | 
                        System.Reflection.BindingFlags.Instance);
                    
                    if (controlAttributesField != null)
                    {
                        var controlAttributes = controlAttributesField.GetValue(VirtualGrid) as Dictionary<string, DwObjectAttributesBase>;
                        
                        if (controlAttributes != null && controlAttributes.Count > 0)
                        {
                            Console.WriteLine($"Direct call to Write: Processing {controlAttributes.Count} attributes");
                            
                            // Process each band's rows with the attributes
                            foreach (var band in VirtualGrid.BandRows)
                            {
                                if (band.Rows.Count > 0)
                                {
                                    var exportedCells = WriteRows(band.Rows, controlAttributes);
                                    if (exportedCells == null || exportedCells.Count == 0)
                                    {
                                        Console.WriteLine($"Warning: No cells were written for band {band.Name}");
                                    }
                                    else
                                    {
                                        Console.WriteLine($"Wrote {exportedCells.Count} cells for band {band.Name}");
                                    }
                                }
                            }
                        }
                        else
                        {
                            Console.WriteLine("Warning: No control attributes found in VirtualGrid for rendering");
                        }
                    }
                    else
                    {
                        Console.WriteLine("Warning: Could not access _controlAttributes field in VirtualGrid");
                    }
                }
                else if (calledFromWriteEntireGrid)
                {
                    Console.WriteLine("Called from WriteEntireGrid, skipping row processing");
                }
                
                // Verify we have written data to the sheet
                if (_sheet.PhysicalNumberOfRows == 0)
                {
                    Console.WriteLine("Warning: Sheet has no rows after processing");
                    
                    // Create at least one row to ensure the sheet exists in the output
                    var emptyRow = _sheet.CreateRow(0);
                    emptyRow.CreateCell(0);
                    Console.WriteLine("Created an empty row/cell to ensure sheet is present in the workbook");
                }
                
                Console.WriteLine("Writing workbook to file...");
                
                // Different file handling approach to avoid the stream closing issues
                try
                {
                    // Clean up our file stream if it's still open
                    if (_destinationFile != null)
                    {
                        Console.WriteLine("Closing existing file stream");
                        _destinationFile.Close();
                        _destinationFile.Dispose();
                        _destinationFile = null;
                    }
                    
                    // Use a completely new stream just for writing the final output
                    Console.WriteLine($"Creating new file stream for final output: {_path}");
                    using (var fs = new FileStream(_path, FileMode.Create, FileAccess.Write))
                    {
                        _workbook.Write(fs);
                        Console.WriteLine("Workbook written successfully");
                    }
                    Console.WriteLine($"File written: {_path}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error writing file: {ex.Message}");
                    throw;
                }
                
                _closed = true;
                
                // Verify the file was created properly
                if (File.Exists(_path))
                {
                    var fileInfo = new FileInfo(_path);
                    Console.WriteLine($"Created file size: {fileInfo.Length} bytes");
                    
                    // Additional verification step - try to open the file to ensure it's valid
                    try
                    {
                        using var testStream = new FileStream(_path, FileMode.Open, FileAccess.Read);
                        using var testWorkbook = new XSSFWorkbook(testStream);
                        Console.WriteLine($"Verification: File contains {testWorkbook.NumberOfSheets} sheet(s)");
                    }
                    catch (Exception vex)
                    {
                        Console.WriteLine($"Warning: Created file may be corrupted: {vex.Message}");
                    }
                }
                else
                {
                    Console.WriteLine("Warning: File does not exist after writing");
                }
                
                return true;
            }
            catch (IOException e)
            {
                error = $"IO Error: {e.Message}";
                Console.WriteLine($"IO Exception: {e}");
                return false;
            }
            catch (Exception e)
            {
                error = $"Unexpected error: {e.Message}";
                Console.WriteLine($"Exception: {e}");
                return false;
            }
            finally
            {
                try
                {
                    // Clean up resources
                    if (_workbook != null && _closed)
                    {
                        _workbook.Close();
                        _workbook.Dispose();
                        _workbook = null;
                    }
                    
                    if (_destinationFile != null)
                    {
                        _destinationFile.Close();
                        _destinationFile.Dispose();
                        _destinationFile = null;
                    }
                    
                    if (_closed)
                    {
                        _sheet = null;
                        _drawingPatriarch = null;
                        
                        if (_resizedColumns != null)
                        {
                            _resizedColumns.Clear();
                            _resizedColumns = null;
                        }
                        
                        if (_pictureCache != null)
                        {
                            _pictureCache.Clear();
                            _pictureCache = null;
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error in cleanup: {ex.Message}");
                }
            }
        }
    }
}
