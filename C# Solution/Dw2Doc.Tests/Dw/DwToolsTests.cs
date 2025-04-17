using Microsoft.VisualStudio.TestTools.UnitTesting;
using Appeon.DotnetDemo.Dw2Doc.Dw;
using Appeon.DotnetDemo.Dw2Doc.Common.Enums;
using Appeon.DotnetDemo.Dw2Doc.Common.DwObjects;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Appeon.DotnetDemo.Dw2Doc.Tests.Dw
{
    [TestClass]
    public class DwToolsTests
    {
        [TestMethod]
        public void GetBands_ValidSyntax_ReturnsBandsCorrectly()
        {
            // Arrange
            string dwSyntax = "header(height=100)\r\ndetail(height=200)\r\nsummary(height=150)\r\nfooter(height=120)\r\ngroup(level=1 header.height=80 trailer.height=60)";
            double yrate = 1.0;
            string[] excludedBands = Array.Empty<string>();

            // Act
            var result = DwTools.GetBands(dwSyntax, yrate, excludedBands);

            // Assert
            Assert.IsNotNull(result);
            Assert.IsTrue(result.Any(b => b.Name == "header"));
            Assert.IsTrue(result.Any(b => b.Name == "detail"));
            Assert.IsTrue(result.Any(b => b.Name == "summary"));
            Assert.IsTrue(result.Any(b => b.Name == "footer"));
            Assert.IsTrue(result.Any(b => b.Name == "header.1"));
            Assert.IsTrue(result.Any(b => b.Name == "trailer.1"));
        }

        [TestMethod]
        public void GetBands_WithExcludedBands_ExcludesBandsCorrectly()
        {
            // Arrange
            string dwSyntax = "header(height=100)\r\ndetail(height=200)\r\nsummary(height=150)\r\nfooter(height=120)";
            double yrate = 1.0;
            string[] excludedBands = new[] { "footer" };

            // Act
            var result = DwTools.GetBands(dwSyntax, yrate, excludedBands);

            // Assert
            Assert.IsNotNull(result);
            Assert.IsTrue(result.Any(b => b.Name == "header"));
            Assert.IsTrue(result.Any(b => b.Name == "detail"));
            Assert.IsTrue(result.Any(b => b.Name == "summary"));
            Assert.IsFalse(result.Any(b => b.Name == "footer"));
        }

        [TestMethod]
        [ExpectedException(typeof(ArgumentException))]
        public void GetBands_InvalidSyntax_ThrowsArgumentException()
        {
            // Arrange
            string dwSyntax = "invalid syntax";
            double yrate = 1.0;
            string[] excludedBands = Array.Empty<string>();

            // Act
            DwTools.GetBands(dwSyntax, yrate, excludedBands);

            // Assert is handled by ExpectedException
        }

        [TestMethod]
        public void GetDwType_ValidType_ReturnsCorrectType()
        {
            // Arrange
            string typeIndex = "0"; // Default type

            // Act
            int result = DwTools.GetDwType(typeIndex, out DwType type, out string? error);

            // Assert
            Assert.AreEqual(0, result);
            Assert.AreEqual(DwType.Default, type);
            Assert.IsNull(error);
        }

        [TestMethod]
        public void GetDwType_GridType_ReturnsCorrectType()
        {
            // Arrange
            string typeIndex = "1"; // Grid type

            // Act
            int result = DwTools.GetDwType(typeIndex, out DwType type, out string? error);

            // Assert
            Assert.AreEqual(0, result);
            Assert.AreEqual(DwType.Grid, type);
            Assert.IsNull(error);
        }

        [TestMethod]
        public void GetDwType_InvalidType_ReturnsError()
        {
            // Arrange
            string typeIndex = "invalid";

            // Act
            int result = DwTools.GetDwType(typeIndex, out DwType type, out string? error);

            // Assert
            Assert.AreEqual(-1, result);
            Assert.IsNotNull(error);
        }

        [TestMethod]
        public void GetAlignment_ReturnsCorrectAlignment()
        {
            // Test all alignment values
            Assert.AreEqual(Alignment.Left, DwTools.GetAlignment(0));
            Assert.AreEqual(Alignment.Right, DwTools.GetAlignment(1));
            Assert.AreEqual(Alignment.Center, DwTools.GetAlignment(2));
            Assert.AreEqual(Alignment.Justify, DwTools.GetAlignment(3));
        }
    }
} 