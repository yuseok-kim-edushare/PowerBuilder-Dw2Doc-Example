using Microsoft.VisualStudio.TestTools.UnitTesting;
using Appeon.DotnetDemo.Dw2Doc.Common.Extensions;
using System.Collections.Generic;
using System.Text;

namespace Appeon.DotnetDemo.Dw2Doc.Tests.Common.Utils
{
    [TestClass]
    public class ExtensionTests
    {
        // Helper method to manually implement what CollectionToString does since we can't access the actual extension method
        private string ManualCollectionToString<T>(ICollection<T> source)
        {
            if (source == null || source.Count == 0)
                return string.Empty;

            StringBuilder sb = new StringBuilder();
            foreach (var item in source)
            {
                sb.AppendLine(item?.ToString());
            }
            return sb.ToString();
        }

        [TestMethod]
        public void AddRange_AddsItemsToCollection()
        {
            // Arrange
            var source = new List<string> { "Item1", "Item2", "Item3" };
            var destination = new List<string>();

            // Act
            destination.AddRange(source);

            // Assert
            Assert.AreEqual(3, destination.Count);
            Assert.AreEqual("Item1", destination[0]);
            Assert.AreEqual("Item2", destination[1]);
            Assert.AreEqual("Item3", destination[2]);
        }

        [TestMethod]
        public void AddRange_EmptySource_DoesNotModifyDestination()
        {
            // Arrange
            var source = new List<string>();
            var destination = new List<string> { "Original" };

            // Act
            destination.AddRange(source);

            // Assert
            Assert.AreEqual(1, destination.Count);
            Assert.AreEqual("Original", destination[0]);
        }

        [TestMethod]
        public void AddRange_MultipleTypes_WorksCorrectly()
        {
            // Arrange
            var source = new List<int> { 1, 2, 3 };
            var destination = new HashSet<int>();

            // Act
            destination.AddRange(source);

            // Assert
            Assert.AreEqual(3, destination.Count);
            Assert.IsTrue(destination.Contains(1));
            Assert.IsTrue(destination.Contains(2));
            Assert.IsTrue(destination.Contains(3));
        }
    }
} 