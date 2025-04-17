using Microsoft.VisualStudio.TestTools.UnitTesting;
using Appeon.DotnetDemo.Dw2Doc.DwData.DataSet;
using Appeon.DotnetDemo.Dw2Doc.Common.DwObjects.DwObjectAttributes;
using System.Collections.Generic;
using System.Linq;

namespace Appeon.DotnetDemo.Dw2Doc.Tests.DwData.DataSet
{
    [TestClass]
    public class DataSetBuilderTests
    {
        private DataSetBuilder _builder = null!;
         /*Using null! tells the compiler
          that we're aware these fields appear to be null at declaration
           but will be properly initialized before use*/
        private MockAttribute _mockAttribute = null!;

        [TestInitialize]
        public void Initialize()
        {
            _builder = new DataSetBuilder();
            _mockAttribute = new MockAttribute();
        }

        [TestMethod]
        public void AddControlAttribute_AddsAttributeCorrectly()
        {
            // Arrange
            string controlName = "testControl";

            // Act
            _builder.AddControlAttribute(controlName, _mockAttribute);
            var result = _builder.Build();

            // Assert
            Assert.IsTrue(result.ContainsKey(controlName));
            Assert.AreEqual(_mockAttribute, result[controlName]);
        }

        [TestMethod]
        public void AddControlAttribute_OverwritesExistingAttribute()
        {
            // Arrange
            string controlName = "testControl";
            var firstAttribute = new MockAttribute { Id = 1 };
            var secondAttribute = new MockAttribute { Id = 2 };

            // Act
            _builder.AddControlAttribute(controlName, firstAttribute);
            _builder.AddControlAttribute(controlName, secondAttribute);
            var result = _builder.Build();

            // Assert
            Assert.IsTrue(result.ContainsKey(controlName));
            Assert.AreEqual(secondAttribute, result[controlName]);
            Assert.AreEqual(2, ((MockAttribute)result[controlName]).Id);
        }

        [TestMethod]
        public void Clear_RemovesAllAttributes()
        {
            // Arrange
            _builder.AddControlAttribute("control1", new MockAttribute());
            _builder.AddControlAttribute("control2", new MockAttribute());

            // Act
            _builder.Clear();
            var result = _builder.Build();

            // Assert
            Assert.AreEqual(0, result.Count);
        }

        [TestMethod]
        public void Build_ReturnsCorrectDictionary()
        {
            // Arrange
            _builder.AddControlAttribute("control1", new MockAttribute { Id = 1 });
            _builder.AddControlAttribute("control2", new MockAttribute { Id = 2 });

            // Act
            var result = _builder.Build();

            // Assert
            Assert.AreEqual(2, result.Count);
            Assert.IsTrue(result.ContainsKey("control1"));
            Assert.IsTrue(result.ContainsKey("control2"));
            Assert.AreEqual(1, ((MockAttribute)result["control1"]).Id);
            Assert.AreEqual(2, ((MockAttribute)result["control2"]).Id);
        }

        // Helper class for testing
        private class MockAttribute : DwObjectAttributesBase
        {
            public int Id { get; set; }
        }
    }
} 