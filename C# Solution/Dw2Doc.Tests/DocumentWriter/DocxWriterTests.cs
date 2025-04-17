using Microsoft.VisualStudio.TestTools.UnitTesting;
using Appeon.DotnetDemo.DocumentWriter;
using NPOI.XWPF.UserModel;
using System;
using System.IO;

namespace Appeon.DotnetDemo.Dw2Doc.Tests.DocumentWriter
{
    [TestClass]
    public class DocxWriterTests
    {
        private readonly string _testOutputPath = Path.Combine(Path.GetTempPath(), "DocxWriterTest.docx");

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
        public void CreateDocument_Success_ReturnsDocument()
        {
            // Act
            int result = DocxWriter.CreateDocument(out XWPFDocument? document, out string? error);

            // Assert
            Assert.AreEqual(1, result);
            Assert.IsNull(error);
            Assert.IsNotNull(document);
        }

        [TestMethod]
        public void WriteToDocument_SingleValue_AppendsCorrectly()
        {
            // Arrange
            DocxWriter.CreateDocument(out XWPFDocument? document, out _);
            Assert.IsNotNull(document);

            // Act
            int result = DocxWriter.WriteToDocument(
                ref document,
                "TestColumn",
                "TestValue",
                "NULL",
                true,
                out string? error);

            // Assert
            Assert.AreEqual(1, result);
            Assert.IsNull(error);
            Assert.AreEqual(1, document.Paragraphs.Count);
            Assert.IsTrue(document.Paragraphs[0].Text.Contains("TestColumn: TestValue"));
        }

        [TestMethod]
        public void WriteToDocument_MultipleValues_AppendsCorrectly()
        {
            // Arrange
            DocxWriter.CreateDocument(out XWPFDocument? document, out _);
            Assert.IsNotNull(document);
            
            string[] data = new[] { "Name=John", "Age=30", "City=New York" };

            // Act
            int result = DocxWriter.WriteToDocument(
                document,
                data,
                "=",
                "NULL",
                out string? error);

            // Assert
            Assert.AreEqual(1, result);
            Assert.IsNull(error);
            Assert.AreEqual(1, document.Paragraphs.Count);
            string paragraphText = document.Paragraphs[0].Text;
            Assert.IsTrue(paragraphText.Contains("Name: John"));
            Assert.IsTrue(paragraphText.Contains("Age: 30"));
            Assert.IsTrue(paragraphText.Contains("City: New York"));
        }

        [TestMethod]
        public void SaveDocument_ValidPath_SavesSuccessfully()
        {
            // Arrange
            DocxWriter.CreateDocument(out XWPFDocument? document, out _);
            Assert.IsNotNull(document);
            DocxWriter.WriteToDocument(
                ref document,
                "TestColumn",
                "TestValue",
                "NULL",
                true,
                out _);

            // Act
            int result = DocxWriter.SaveDocument(document, _testOutputPath, out string? error);

            // Assert
            Assert.AreEqual(1, result);
            Assert.IsNull(error);
            Assert.IsTrue(File.Exists(_testOutputPath));
        }

        [TestMethod]
        public void SaveDocument_InvalidPath_ReturnsError()
        {
            // Arrange
            DocxWriter.CreateDocument(out XWPFDocument? document, out _);
            Assert.IsNotNull(document);
            string invalidPath = Path.Combine("Z:\\NonExistentFolder", "test.docx");

            // Act
            int result = DocxWriter.SaveDocument(document, invalidPath, out string? error);

            // Assert
            Assert.AreEqual(-1, result);
            Assert.IsNotNull(error);
        }
    }
} 