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
        private int _startRowOffset;
        private int _startColumnOffset;
        private int _currentRowOffset;
        private XSSFWorkbook? _workbook;
        private XSSFSheet? _sheet;
        private XSSFDrawing? _drawingPatriarch;
        private ISet<ColumnDefinition>? _resizedColumns;
        private IDictionary<string, int>? _pictureCache;
        private bool _writerInitialized = false;
        private bool _appendingToSheet = false;
        private string _path;
        private FileStream? _destinationFile;
        private bool _closed = false;
        private Dictionary<string, VirtualCell?> rangedCells;

        internal VirtualGridXlsxWriter(VirtualGrid grid, string path)
            : base(grid)
        {
            _workbook = new XSSFWorkbook();
            _sheet = (XSSFSheet)_workbook.CreateSheet();
            _drawingPatriarch = (XSSFDrawing)_sheet.CreateDrawingPatriarch();
            _resizedColumns = new HashSet<ColumnDefinition>();

            _resizedColumns.AddRange(VirtualGrid.Columns);
            _pictureCache = new Dictionary<string, int>();
            _path = path;
            rangedCells = new();
            LockFile(_path);
        }

        internal VirtualGridXlsxWriter(VirtualGrid grid, string path, string sheetname) : base(grid)
        {
            _path = path;
            LockFile(path);

            _workbook = new XSSFWorkbook(new FileStream(path, FileMode.Open, FileAccess.Read));

            if (_workbook.GetSheet(sheetname) is not null)
            {
                throw new ArgumentException("Specified sheet name already exists", nameof(sheetname));
            }

            _sheet = (XSSFSheet)_workbook.CreateSheet(sheetname);
            _workbook.SetActiveSheet(_workbook.GetSheetIndex(_sheet));

            _drawingPatriarch = (XSSFDrawing)_sheet.CreateDrawingPatriarch();
            _resizedColumns = new HashSet<ColumnDefinition>();

            _resizedColumns.AddRange(VirtualGrid.Columns);
            _pictureCache = new Dictionary<string, int>();

            _appendingToSheet = true;
        }

        [MemberNotNull(nameof(_destinationFile))]
        private void LockFile(string _path)
        {
            _destinationFile = File.Exists(_path)
                    ? new FileStream(_path, FileMode.Truncate, FileAccess.Write)
                    : new FileStream(_path, FileMode.Create, FileAccess.Write);
        }

        private void InitWriter()
        {
            if (_writerInitialized) return;

            var row = _sheet.CreateRow(_startRowOffset);
            foreach (var column in _resizedColumns)
            {
                row.CreateCell(column.IndexOffset + _startColumnOffset);
                _sheet.SetColumnWidth(column.IndexOffset + _startColumnOffset, (int)UnitConversion.PixelsToColumnWidth(column.Size));
            }

            _sheet.RemoveRow(row);

            _writerInitialized = true;
        }

        protected override IList<ExportedCellBase>? WriteRows(IList<RowDefinition> rows, IDictionary<string, DwObjectAttributesBase>? data)
        {
            InitWriter();
            if (data is null)
                return null;

            try
            {
                if (_closed)
                {
                    throw new InvalidOperationException("This writer has already been closed");
                }

                var exportedCells = new List<ExportedCellBase>();
                ExportedCellBase? exportedCell = null;

                foreach (var row in rows)
                {
                    var xRow = _sheet.CreateRow(_startRowOffset + _currentRowOffset);

                    checked
                    {
                        xRow.Height = (short)(row.Size).PixelsToTwips(Windows.ScreenTools.MeasureDirection.Vertical);
                        //xRow.Height = (short)row.Size;
                    }

                    ICell? previousCell = null;
                    int lastOccupiedColumn = 0;
                    foreach (var cell in row.Objects.Concat(row.FloatingObjects))
                    {
                        var attribute = data[cell.Object.Name];

                        if (!attribute.IsVisible)
                        {
                            continue;
                        }
                        var renderer = RendererLocator.Find(attribute.GetType())
                            ?? throw new InvalidOperationException($"Could not find renderer for attribute [{attribute.GetType().FullName}]");

                        if (renderer is XlsxPictureRenderer pictureRenderer)
                        {
                            pictureRenderer.SetPictureCache(_pictureCache);
                        }

                        if (cell.OwningColumn is not null && !cell.Object.Floating)
                        { // cell is not floating
                            var xCell = xRow.CreateCell(cell.OwningColumn.IndexOffset + _startColumnOffset);

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
                                throw new InvalidOperationException("Non-floating cell in FloatingObjects list");
                            previousCell = null;
                            exportedCell = renderer.Render(_sheet, floatingCell,
                                data[cell.Object.Name],
                                (_startColumnOffset + floatingCell.Offset.StartColumn.IndexOffset,
                                    _currentRowOffset,
                                    _drawingPatriarch));

                            if (exportedCell is not null)
                            {
                                exportedCells.Add(exportedCell);
                            }
                        }
                    }


                    int unoccupiedTrailingColumns = 0;
                    if (row.IsFiller
                        && VirtualGrid.Columns.Count > 1
                        && row.Objects.Count == 0
                        && row.FloatingObjects.Count == 0
                        || (unoccupiedTrailingColumns = VirtualGrid.Columns.Count - lastOccupiedColumn) > 2)
                    {
                        var cellRange = new CellRangeAddress(
                            _currentRowOffset + _startRowOffset,
                            _currentRowOffset + _startRowOffset,
                            _startColumnOffset + unoccupiedTrailingColumns > 2 ? (lastOccupiedColumn + 1) : 0,
                            _startColumnOffset + VirtualGrid.Columns.Count - 1);
                        rangedCells[cellRange.FormatAsString()] = null;
                        _sheet.AddMergedRegion(cellRange);
                    }

                    ++_currentRowOffset;
                }

                return exportedCells;
            }
            catch (Exception ex)
            {
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

            if (_path is null)
            {
                error = "No file specified";
                return false;
            }
            if (_closed)
            {
                error = "This writer has already been closed";
                return false;
            }
            try
            {
                _workbook.Write(_destinationFile);

                _workbook.Close();
                _closed = true;
            }
            catch (IOException e)
            {
                error = e.Message;
                return false;
            }
            finally
            {
                _workbook.Dispose();
                _workbook = null;
                _sheet = null;
                _drawingPatriarch = null;
                _resizedColumns.Clear();
                _resizedColumns = null;
                _pictureCache.Clear();
                _pictureCache = null;
            }

            return true;
        }
    }
}
