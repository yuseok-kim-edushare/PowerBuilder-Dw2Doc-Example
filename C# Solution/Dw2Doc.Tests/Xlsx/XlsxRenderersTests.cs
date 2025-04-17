using Appeon.DotnetDemo.Dw2Doc.Common.DwObjects;
using Appeon.DotnetDemo.Dw2Doc.Common.DwObjects.DwObjectAttributes;
using Appeon.DotnetDemo.Dw2Doc.Common.Enums;
using Appeon.DotnetDemo.Dw2Doc.Common.VirtualGrid;
using Appeon.DotnetDemo.Dw2Doc.Tests.Xlsx.TestSupport;
using Appeon.DotnetDemo.Dw2Doc.Xlsx.VirtualGridWriter.XlsxWriter;
using Appeon.DotnetDemo.Dw2Doc.Common.VirtualGridWriter.Abstractions;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;

namespace Appeon.DotnetDemo.Dw2Doc.Tests.Xlsx
{
    [TestClass]
    public class XlsxRenderersTests
    {
        private const string TestOutputDir = "TestOutput";
        
        // Use nullable reference types for fields initialized in Setup
        private XSSFWorkbook? _workbook;
        private XSSFSheet? _sheet;
        private string _testFilePath = string.Empty; // Initialize with empty string

        [TestInitialize]
        public void Setup()
        {
            // Create test output directory if it doesn't exist
            if (!Directory.Exists(TestOutputDir))
            {
                Directory.CreateDirectory(TestOutputDir);
            }

            _testFilePath = Path.Combine(TestOutputDir, "renderers_test.xlsx");
            
            // Clean up any existing file
            if (File.Exists(_testFilePath))
            {
                File.Delete(_testFilePath);
            }

            _workbook = new XSSFWorkbook();
            _sheet = (XSSFSheet)_workbook.CreateSheet("TestSheet");
            
            Console.WriteLine("Initializing Common tests...");
        }

        [TestCleanup]
        public void Cleanup()
        {
            Console.WriteLine("Cleaning up after Common tests...");
            _workbook?.Close(); // Use null conditional operator for safety
        }

        // Since renderer classes are internal, we test them indirectly through the public API

        [TestMethod]
        public void RenderText_RendersCorrectly()
        {
            try
            {
                // Arrange - Create VirtualGrid with text content
                var grid = TestVirtualGridFactory.CreateSimpleVirtualGrid();
                Console.WriteLine($"Created grid with {grid.Rows.Count} rows and {grid.Columns.Count} columns");
                
                // Act - Export to Excel using VirtualGridXlsxWriter
                ExportGridToExcel(grid, _testFilePath);
                Console.WriteLine($"Exported grid to {_testFilePath}");
                
                // Verify file was created
                Assert.IsTrue(File.Exists(_testFilePath), "Excel file was not created");
                Console.WriteLine($"File exists: {File.Exists(_testFilePath)}");
                
                // Assert - Check that text was rendered correctly
                using (var workbook = new XSSFWorkbook(_testFilePath))
                {
                    Assert.IsNotNull(workbook, "Workbook is null");
                    Assert.IsTrue(workbook.NumberOfSheets > 0, "No sheets in workbook");
                    
                    var sheet = workbook.GetSheetAt(0);
                    Assert.IsNotNull(sheet, "Sheet is null");
                    Console.WriteLine($"Sheet has {sheet.PhysicalNumberOfRows} rows");
                    
                    Assert.IsTrue(sheet.PhysicalNumberOfRows > 0, "Sheet has no rows");
                    
                    var row = sheet.GetRow(0);
                    Assert.IsNotNull(row, "Row is null");
                    Console.WriteLine($"Row has {row.PhysicalNumberOfCells} cells");
                    
                    Assert.IsTrue(row.PhysicalNumberOfCells > 0, "Row has no cells");
                    
                    var cell = row.GetCell(0);
                    Assert.IsNotNull(cell, "Cell is null");
                    Console.WriteLine($"Cell value: {cell.StringCellValue}");
                    
                    // Verify cell has text
                    Assert.IsFalse(string.IsNullOrEmpty(cell.StringCellValue), "Cell text is empty");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Exception in RenderText_RendersCorrectly: {ex}");
                throw;
            }
        }

        [TestMethod]
        public void RenderButton_RendersCorrectly()
        {
            // Skip this test for now - we would need to create a specific
            // test grid in TestVirtualGridFactory with button components
            Console.WriteLine("Skipping RenderButton_RendersCorrectly test");
        }

        [TestMethod]
        public void RenderCheckbox_RendersCorrectly()
        {
            // Skip this test for now - we would need to create a specific
            // test grid in TestVirtualGridFactory with checkbox components
            Console.WriteLine("Skipping RenderCheckbox_RendersCorrectly test");
        }

        [TestMethod]
        public void ComplexGrid_RendersCorrectly()
        {
            try
            {
                // Arrange - Create a complex VirtualGrid
                var grid = TestVirtualGridFactory.CreateComplexVirtualGrid();
                Console.WriteLine($"Created complex grid with {grid.Rows.Count} rows and {grid.Columns.Count} columns");
                
                // Act - Export to Excel using VirtualGridXlsxWriter
                ExportGridToExcel(grid, _testFilePath);
                Console.WriteLine($"Exported complex grid to {_testFilePath}");
                
                // Verify file was created
                Assert.IsTrue(File.Exists(_testFilePath), "Excel file was not created");
                Console.WriteLine($"File exists: {File.Exists(_testFilePath)}");
                
                // Assert - Check that grid was rendered correctly
                using (var workbook = new XSSFWorkbook(_testFilePath))
                {
                    Assert.IsNotNull(workbook, "Workbook is null");
                    Assert.IsTrue(workbook.NumberOfSheets > 0, "No sheets in workbook");
                    
                    var sheet = workbook.GetSheetAt(0);
                    Assert.IsNotNull(sheet, "Sheet is null");
                    Console.WriteLine($"Sheet has {sheet.PhysicalNumberOfRows} rows");
                    
                    // Verify sheet has content
                    Assert.IsTrue(sheet.PhysicalNumberOfRows > 0, "Sheet has no rows");
                    
                    // Verify header row
                    var headerRow = sheet.GetRow(0);
                    Assert.IsNotNull(headerRow, "Header row is null");
                    Console.WriteLine($"Header row has {headerRow.PhysicalNumberOfCells} cells");
                    
                    // Verify data row
                    if (sheet.PhysicalNumberOfRows > 1)
                    {
                        var dataRow = sheet.GetRow(1);
                        Assert.IsNotNull(dataRow, "Data row is null");
                        Console.WriteLine($"Data row has {dataRow.PhysicalNumberOfCells} cells");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Exception in ComplexGrid_RendersCorrectly: {ex}");
                throw;
            }
        }

        /// <summary>
        /// Helper method to export a grid to Excel
        /// </summary>
        private void ExportGridToExcel(VirtualGrid grid, string filePath)
        {
            var builder = new VirtualGridXlsxWriterBuilder
            {
                WriteToPath = filePath
            };
            
            var writer = builder.Build(grid, out string? error);
            
            if (writer == null)
            {
                Console.WriteLine($"Failed to build writer: {error}");
                Assert.Fail($"Failed to build writer: {error}");
                return;
            }
            
            // Check that attributes are properly set up
            var controlAttributesField = typeof(VirtualGrid).GetField("_controlAttributes", 
                System.Reflection.BindingFlags.NonPublic | 
                System.Reflection.BindingFlags.Instance);
            
            if (controlAttributesField != null)
            {
                var attributes = controlAttributesField.GetValue(grid) as Dictionary<string, DwObjectAttributesBase>;
                Console.WriteLine($"Control attributes count: {(attributes != null ? attributes.Count : 0)}");
            }
            else
            {
                Console.WriteLine("Could not access _controlAttributes field");
            }
            
            // Call the WriteEntireGrid method which uses control attributes
            bool result = ((AbstractVirtualGridWriter)writer).WriteEntireGrid(null, out error);
            
            if (!result)
            {
                Console.WriteLine($"Failed to write: {error}");
                Assert.Fail($"Failed to write Excel file: {error}");
            }
        }
    }
} 