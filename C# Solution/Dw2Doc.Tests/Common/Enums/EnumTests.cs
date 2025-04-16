using Microsoft.VisualStudio.TestTools.UnitTesting;
using Appeon.DotnetDemo.Dw2Doc.Common.Enums;
using System;

namespace Appeon.DotnetDemo.Dw2Doc.Tests.Common.Enums
{
    [TestClass]
    public class EnumTests
    {
        [TestMethod]
        public void BandType_EnumValuesExist()
        {
            // Assert
            Assert.IsTrue(Enum.IsDefined(typeof(BandType), BandType.Header));
            Assert.IsTrue(Enum.IsDefined(typeof(BandType), BandType.Trailer));
        }

        [TestMethod]
        public void DwType_EnumValuesExist()
        {
            // Assert
            Assert.IsTrue(Enum.IsDefined(typeof(DwType), DwType.Grid));
            Assert.IsTrue(Enum.IsDefined(typeof(DwType), DwType.Default));
            Assert.IsTrue(Enum.IsDefined(typeof(DwType), DwType.Label));
            Assert.IsTrue(Enum.IsDefined(typeof(DwType), DwType.Composite));
            Assert.IsTrue(Enum.IsDefined(typeof(DwType), DwType.Graph));
            Assert.IsTrue(Enum.IsDefined(typeof(DwType), DwType.Crosstab));
            Assert.IsTrue(Enum.IsDefined(typeof(DwType), DwType.OLE));
            Assert.IsTrue(Enum.IsDefined(typeof(DwType), DwType.RichText));
            Assert.IsTrue(Enum.IsDefined(typeof(DwType), DwType.TreeView));
            Assert.IsTrue(Enum.IsDefined(typeof(DwType), DwType.TreeViewWithGrid));
        }

        [TestMethod]
        public void Alignment_EnumValuesExist()
        {
            // Assert
            Assert.IsTrue(Enum.IsDefined(typeof(Alignment), Alignment.Left));
            Assert.IsTrue(Enum.IsDefined(typeof(Alignment), Alignment.Right));
            Assert.IsTrue(Enum.IsDefined(typeof(Alignment), Alignment.Center));
            Assert.IsTrue(Enum.IsDefined(typeof(Alignment), Alignment.Justify));
        }

        [TestMethod]
        public void LineStyle_EnumValuesExist()
        {
            // Assert
            Assert.IsTrue(Enum.IsDefined(typeof(LineStyle), LineStyle.Solid));
            Assert.IsTrue(Enum.IsDefined(typeof(LineStyle), LineStyle.Dash));
            Assert.IsTrue(Enum.IsDefined(typeof(LineStyle), LineStyle.DashDot));
            Assert.IsTrue(Enum.IsDefined(typeof(LineStyle), LineStyle.DashDotDot));
            Assert.IsTrue(Enum.IsDefined(typeof(LineStyle), LineStyle.Dotted));
            Assert.IsTrue(Enum.IsDefined(typeof(LineStyle), LineStyle.NoLine));
        }
    }
} 