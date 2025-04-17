using Appeon.DotnetDemo.Dw2Doc.Common.DwObjects.DwObjectAttributes;
using Appeon.DotnetDemo.Dw2Doc.Common.VirtualGrid;
using Appeon.DotnetDemo.Dw2Doc.Common.VirtualGridWriter.Models;
using Appeon.DotnetDemo.Dw2Doc.Xlsx.Models;
using Appeon.DotnetDemo.Dw2Doc.XlsxTester;
using Appeon.DotnetDemo.Dw2Doc.XlsxTester.AttributeTester;
using Appeon.DotnetDemo.Dw2Doc.XlsxTester.Models;
using Appeon.DotnetDemo.Dw2Doc.Tests.Xlsx.TestSupport;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace Appeon.DotnetDemo.Dw2Doc.Tests.XlsxTester
{
    [TestClass]
    public class XlsxTesterTests
    {
        [TestMethod]
        public void TesterForAttribute_Constructor_SetsProperties()
        {
            // Arrange
            var attributeType = typeof(DwTextAttributes);
            var testerType = typeof(TextTester);

            // Act
            var attribute = new TesterForAttribute(attributeType, testerType);

            // Assert
            Assert.AreEqual(attributeType, attribute.AttributeType);
            Assert.AreEqual(testerType, attribute.TesterType);
        }

        [TestMethod]
        public void AttributeTestResult_Constructor_SetsProperties()
        {
            // Arrange
            string objectName = "TestObject";
            string attributeName = "TestAttribute";
            string expectedValue = "ExpectedValue";
            string realValue = "RealValue";

            // Act
            var result = new AttributeTestResult(objectName, attributeName, expectedValue, realValue);

            // Assert
            Assert.AreEqual(objectName, result.ObjectName);
            Assert.AreEqual(attributeName, result.Attribute);
            Assert.AreEqual(expectedValue, result.ExpectedValue);
            Assert.AreEqual(realValue, result.RealValue);
        }

        [TestMethod]
        public void AttributeTestResult_Result_ReturnsTrueWhenValuesMatch()
        {
            // Arrange
            var result = new AttributeTestResult("TestObject", "TestAttribute", "SameValue", "SameValue");

            // Assert
            Assert.IsTrue(result.Result);
        }

        [TestMethod]
        public void AttributeTestResult_Result_ReturnsFalseWhenValuesDontMatch()
        {
            // Arrange
            var result = new AttributeTestResult("TestObject", "TestAttribute", "Value1", "Value2");

            // Assert
            Assert.IsFalse(result.Result);
        }

        [TestMethod]
        public void AttributeTestResult_Result_HandlesSimilarNumberValuesCorrectly()
        {
            // Arrange - Should match even if string representations differ
            var result1 = new AttributeTestResult("TestObject", "TestAttribute", "10", "10.0");
            var result2 = new AttributeTestResult("TestObject", "TestAttribute", "10.00", "10");

            // Assert
            Assert.IsTrue(result1.Result);
            Assert.IsTrue(result2.Result);
        }

        [TestMethod]
        public void AttributeTestResultCollection_Constructor_SetsResults()
        {
            // Arrange
            var results = new List<AttributeTestResult>
            {
                new AttributeTestResult("Object1", "Attr1", "Val1", "Val1"),
                new AttributeTestResult("Object2", "Attr2", "Val2", "Val2")
            };

            // Act
            var collection = new AttributeTestResultCollection(results);

            // Assert
            Assert.AreEqual(results, collection.Results);
            Assert.AreEqual(2, collection.Results.Count);
        }

        [TestMethod]
        public void AttributeTestResultCollection_Get_ReturnsCorrectResult()
        {
            // Arrange
            var results = new List<AttributeTestResult>
            {
                new AttributeTestResult("Object1", "Attr1", "Val1", "Val1"),
                new AttributeTestResult("Object2", "Attr2", "Val2", "Val2")
            };
            var collection = new AttributeTestResultCollection(results);

            // Act
            var result = collection.Get(1);

            // Assert
            Assert.AreEqual("Object2", result.ObjectName);
            Assert.AreEqual("Attr2", result.Attribute);
        }

        [TestMethod]
        public void AttributeTesterLocator_Find_ByType_ReturnsCorrectTester()
        {
            // This test assumes there's at least one tester registered for an attribute type
            // Get all types with TesterForAttribute
            var testTypes = Assembly.GetAssembly(typeof(TesterForAttribute))!
                .GetTypes()
                .Where(t => t.GetCustomAttributes<TesterForAttribute>().Any())
                .ToList();
            
            Assert.IsTrue(testTypes.Count > 0, "No types with TesterForAttribute found to test");

            // Get the first type with TesterForAttribute
            var testType = testTypes.First();
            var testerAttr = testType.GetCustomAttribute<TesterForAttribute>();
            Assert.IsNotNull(testerAttr, "TesterForAttribute not found on test type");

            // Act
            var tester = AttributeTesterLocator.Find(testerAttr!.AttributeType);

            // Assert
            Assert.IsNotNull(tester, "Tester should not be null");
            Assert.IsInstanceOfType(tester, testerAttr.TesterType);
        }

        [TestMethod]
        public void AttributeTesterLocator_Find_ByExportedCell_ReturnsCorrectTester()
        {
            // Create a test cell with a specific attribute type
            var attributeType = typeof(DwTextAttributes);
            var virtualGrid = TestVirtualGridFactory.CreateSimpleVirtualGrid();

            // Get a cell from the grid
            var cell = virtualGrid.CellRepository.Cells.First();
            var textAttributes = new DwTextAttributes();
            
            // Create a wrapper to simulate an exported cell with the right attribute type
            var exportedCell = new SimpleExportedCell(cell, textAttributes);

            // Act
            var tester = AttributeTesterLocator.Find(exportedCell);

            // Assert
            Assert.IsNotNull(tester, "Tester should not be null");
            Assert.IsInstanceOfType(tester, typeof(TextTester));
        }
    }

    // Simple implementation of ExportedCellBase for testing
    internal class SimpleExportedCell : ExportedCellBase
    {
        public SimpleExportedCell(VirtualCell cell, DwObjectAttributesBase attribute)
            : base(cell, attribute)
        {
        }
    }
} 