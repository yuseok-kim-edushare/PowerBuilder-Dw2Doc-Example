using Appeon.DotnetDemo.Dw2Doc.Common.DwObjects.DwObjectAttributes;
using Appeon.DotnetDemo.Dw2Doc.Common.Enums;
using Appeon.DotnetDemo.Dw2Doc.Common.VirtualGrid;
using Appeon.DotnetDemo.Dw2Doc.Common.VirtualGridWriter.Models;
using Appeon.DotnetDemo.Dw2Doc.Docx.VirtualGridWriter.DocxWriter;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;
using System.IO;

namespace Appeon.DotnetDemo.Dw2Doc.Tests.Docx
{
    [TestClass]
    public class VirtualGridDocxWriterTests
    {
        private VirtualGrid? _virtualGrid;
        private string? _testOutputPath;

        [TestInitialize]
        public void Setup()
        {
            // Create a minimal virtual grid for testing using reflection
            // since the constructor and some dependencies are internal
            var rows = new List<RowDefinition>();
            var columns = new List<ColumnDefinition>();
            var bandRows = new List<BandRows>();
            
            // Create cell repository using reflection since it's internal
            var cellRepoType = typeof(VirtualGrid).Assembly.GetType("Appeon.DotnetDemo.Dw2Doc.Common.VirtualGrid.VirtualCellRepository");
            Assert.IsNotNull(cellRepoType, "Could not find VirtualCellRepository type");
            
            var cellRepo = System.Activator.CreateInstance(cellRepoType, System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic, null, null, null);
            Assert.IsNotNull(cellRepo, "Could not create VirtualCellRepository instance");
            
            // Create VirtualGrid using reflection
            var constructor = typeof(VirtualGrid).GetConstructors(System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance)[0];
            _virtualGrid = (VirtualGrid)constructor.Invoke(new object[] { rows, columns, bandRows, cellRepo, DwType.Grid });
            
            // Create output directory if it doesn't exist
            _testOutputPath = Path.Combine(Path.GetTempPath(), "DocxWriterTests");
            Directory.CreateDirectory(_testOutputPath);
        }

        [TestCleanup]
        public void Cleanup()
        {
            // Optional: Clean up test files after tests
            // Directory.Delete(_testOutputPath, true);
        }

        [TestMethod]
        public void Constructor_WithValidVirtualGrid_ShouldCreateInstance()
        {
            // Arrange & Act
            Assert.IsNotNull(_virtualGrid, "VirtualGrid should not be null");
            var writer = new VirtualGridDocxWriter(_virtualGrid);

            // Assert
            Assert.IsNotNull(writer);
        }

        [TestMethod]
        public void Write_WithNullPath_ShouldReturnFalseAndError()
        {
            // Arrange
            Assert.IsNotNull(_virtualGrid, "VirtualGrid should not be null");
            var writer = new VirtualGridDocxWriter(_virtualGrid);

            // Act
            bool result = writer.Write(null, out string? error);

            // Assert
            Assert.IsFalse(result);
            Assert.IsNotNull(error);
            Assert.AreEqual("No file specified", error);
        }

        [TestMethod]
        public void Write_WithValidPath_ShouldCreateDocxFile()
        {
            // Arrange
            Assert.IsNotNull(_virtualGrid, "VirtualGrid should not be null");
            Assert.IsNotNull(_testOutputPath, "TestOutputPath should not be null");
            
            var writer = new VirtualGridDocxWriter(_virtualGrid);
            string filePath = Path.Combine(_testOutputPath, "test_output.docx");

            // Act
            bool result = writer.Write(filePath, out string? error);

            // Assert
            Assert.IsTrue(result);
            Assert.IsNull(error);
            Assert.IsTrue(File.Exists(filePath), "DOCX file should be created");
        }

        [TestMethod]
        public void WriteRows_WithEmptyData_ShouldHandleGracefully()
        {
            // Arrange
            Assert.IsNotNull(_virtualGrid, "VirtualGrid should not be null");
            var writer = new VirtualGridDocxWriter(_virtualGrid);
            var rows = new List<RowDefinition>();
            var data = new Dictionary<string, DwObjectAttributesBase>();

            // Act - We need to use reflection since WriteRows is protected
            var method = typeof(VirtualGridDocxWriter).GetMethod("WriteRows", 
                System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
            Assert.IsNotNull(method, "WriteRows method should not be null");
            
            var result = method.Invoke(writer, new object[] { rows, data }) as IList<ExportedCellBase>;

            // Assert
            Assert.IsNull(result);
        }
    }
} 