using Microsoft.VisualStudio.TestTools.UnitTesting;
using Appeon.DotnetDemo.DocumentWriter;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using System;
using System.IO;

namespace Appeon.DotnetDemo.Dw2Doc.Tests.DocumentWriter
{
    [TestClass]
    public class XlsxWriterTests
    {
        private readonly string _testOutputPath = Path.Combine(Path.GetTempPath(), "XlsxWriterTest.xlsx");

        [TestCleanup]
        public void Cleanup()
        {
            // Clean up any test files after each test
            if (File.Exists(_testOutputPath))
            {
                File.Delete(_testOutputPath);
            }
        }

        [TestMethod]
        public void CreateWorkbook_Success_ReturnsWorkbook()
        {
            // Act
            int result = XlsxWriter.CreateWorkbook(out IWorkbook? workbook, out string? error);

            // Assert
            Assert.AreEqual(1, result);
            Assert.IsNull(error);
            Assert.IsNotNull(workbook);
            Assert.AreEqual(1, workbook.NumberOfSheets);
        }

        [TestMethod]
        public void WriteStringArrayToWorkbook_Success_WritesData()
        {
            // Arrange
            XlsxWriter.CreateWorkbook(out IWorkbook? workbook, out _);
            Assert.IsNotNull(workbook);
            string[] data = new[] { "Header1,Header2,Header3", "Value1,Value2,Value3", "Value4,Value5,Value6" };

            // Act
            int result = XlsxWriter.WriteStringArrayToWorkbook(
                ref workbook,
                data,
                ",",
                out string? error);

            // Assert
            Assert.AreEqual(1, result);
            Assert.IsNull(error);
            
            // Verify data was written
            ISheet sheet = workbook.GetSheetAt(0);
            Assert.AreEqual("Header1", sheet.GetRow(0).GetCell(0).StringCellValue);
            Assert.AreEqual("Header2", sheet.GetRow(0).GetCell(1).StringCellValue);
            Assert.AreEqual("Header3", sheet.GetRow(0).GetCell(2).StringCellValue);
            Assert.AreEqual("Value1", sheet.GetRow(1).GetCell(0).StringCellValue);
        }

        [TestMethod]
        public void WriteStringArrayToWorkbook_WithCellAddress_WritesDataAtSpecifiedPosition()
        {
            // Arrange
            XlsxWriter.CreateWorkbook(out IWorkbook? workbook, out _);
            Assert.IsNotNull(workbook);
            string[] data = new[] { "Header1,Header2", "Value1,Value2" };
            CellAddress startAddress = new CellAddress(2, 3); // Row 2, Column 3

            // Act
            int result = XlsxWriter.WriteStringArrayToWorkbook(
                ref workbook,
                data,
                ",",
                startAddress,
                out string? error);

            // Assert
            Assert.AreEqual(1, result);
            Assert.IsNull(error);
            
            // Verify data was written at the correct position
            ISheet sheet = workbook.GetSheetAt(0);
            Assert.AreEqual("Header1", sheet.GetRow(2).GetCell(3).StringCellValue);
            Assert.AreEqual("Header2", sheet.GetRow(2).GetCell(4).StringCellValue);
            Assert.AreEqual("Value1", sheet.GetRow(3).GetCell(3).StringCellValue);
        }

        [TestMethod]
        public void SaveDocument_ValidPath_SavesSuccessfully()
        {
            // Arrange
            XlsxWriter.CreateWorkbook(out IWorkbook? workbook, out _);
            Assert.IsNotNull(workbook);
            string[] data = new[] { "Header1,Header2", "Value1,Value2" };
            XlsxWriter.WriteStringArrayToWorkbook(ref workbook, data, ",", out _);

            // Act
            int result = XlsxWriter.SaveDocument(workbook, _testOutputPath, out string? error);

            // Assert
            Assert.AreEqual(1, result);
            Assert.IsNull(error);
            Assert.IsTrue(File.Exists(_testOutputPath));
        }

        [TestMethod]
        public void SaveDocument_InvalidPath_ReturnsError()
        {
            // Arrange
            XlsxWriter.CreateWorkbook(out IWorkbook? workbook, out _);
            Assert.IsNotNull(workbook);
            string invalidPath = Path.Combine("Z:\\NonExistentFolder", "test.xlsx");

            // Act
            int result = XlsxWriter.SaveDocument(workbook, invalidPath, out string? error);

            // Assert
            Assert.AreEqual(-1, result);
            Assert.IsNotNull(error);
        }
    }
} 