using Microsoft.VisualStudio.TestTools.UnitTesting;
using Appeon.CSharpPbExtensions;
using System.Collections.Generic;

namespace Appeon.DotnetDemo.Dw2Doc.Tests
{
    [TestClass]
    public class ObjectExtensionsTests
    {
        [TestMethod]
        public void ToString_NullObject_ReturnsEmptyString()
        {
            // Arrange
            object? obj = null;
            
            // Act
            string result = ObjectExtensions.ToString(obj);
            
            // Assert
            Assert.AreEqual(string.Empty, result);
        }
        
        [TestMethod]
        public void ToString_StringObject_ReturnsString()
        {
            // Arrange
            string str = "Test String";
            
            // Act
            string result = ObjectExtensions.ToString(str);
            
            // Assert
            Assert.AreEqual(str, result);
        }
        
        [TestMethod]
        public void ToString_IntObject_ReturnsStringRepresentation()
        {
            // Arrange
            int number = 42;
            
            // Act
            string result = ObjectExtensions.ToString(number);
            
            // Assert
            Assert.AreEqual("42", result);
        }
        
        [TestMethod]
        public void ToString_EmptyList_ReturnsEmptyString()
        {
            // Arrange
            var list = new List<string>();
            
            // Act
            string result = ObjectExtensions.ToString(list);
            
            // Assert
            Assert.AreEqual("", result);
        }
        
        [TestMethod]
        public void ToString_ListWithElements_ReturnsCommaSeparatedString()
        {
            // Arrange
            var list = new List<string> { "one", "two", "three" };
            
            // Act
            string result = ObjectExtensions.ToString(list);
            
            // Assert
            Assert.AreEqual("one, two, three", result);
        }
    }

    [TestClass]
    public class StringExtensionsTests
    {
        [TestMethod]
        public void Replace_BasicReplacement_ReturnsExpectedString()
        {
            // Arrange
            string source = "Hello World";
            string pattern = "World";
            string replacement = "Universe";
            
            // Act
            string result = StringExtensions.Replace(source, pattern, replacement);
            
            // Assert
            Assert.AreEqual("Hello Universe", result);
        }
        
        [TestMethod]
        public void Split_BasicSplit_ReturnsSplitArray()
        {
            // Arrange
            string source = "one,two,three";
            string separator = ",";
            string[] expected = { "one", "two", "three" };
            
            // Act
            StringExtensions.Split(source, separator, out string[] result);
            
            // Assert
            CollectionAssert.AreEqual(expected, result);
        }
        
        [TestMethod]
        public void Join_BasicJoin_ReturnsJoinedString()
        {
            // Arrange
            string[] source = { "one", "two", "three" };
            string separator = ", ";
            
            // Act
            string result = StringExtensions.Join(source, separator);
            
            // Assert
            Assert.AreEqual("one, two, three", result);
        }
    }

    [TestClass]
    public class BitOperationsTests
    {
        [TestMethod]
        public void ShiftRightInt_PositiveNumber_ShiftsCorrectly()
        {
            // Arrange
            int number = 16;
            int shift = 2;
            
            // Act
            int result = BitOperations.ShiftRightInt(number, shift);
            
            // Assert
            Assert.AreEqual(4, result);
        }
        
        [TestMethod]
        public void ShiftLeftInt_PositiveNumber_ShiftsCorrectly()
        {
            // Arrange
            int number = 4;
            int shift = 2;
            
            // Act
            int result = BitOperations.ShiftLeftInt(number, shift);
            
            // Assert
            Assert.AreEqual(16, result);
        }
        
        [TestMethod]
        public void BitwiseAndInt_BasicOperation_ReturnsCorrectResult()
        {
            // Arrange
            int number1 = 12; // 1100
            int number2 = 10; // 1010
            
            // Act
            int result = BitOperations.BitwiseAndInt(number1, number2);
            
            // Assert
            Assert.AreEqual(8, result); // 1000
        }
        
        [TestMethod]
        public void BitwiseOrInt_BasicOperation_ReturnsCorrectResult()
        {
            // Arrange
            int number1 = 12; // 1100
            int number2 = 10; // 1010
            
            // Act
            int result = BitOperations.BitwiseOrInt(number1, number2);
            
            // Assert
            Assert.AreEqual(14, result); // 1110
        }
        
        [TestMethod]
        public void BitwiseXorInt_BasicOperation_ReturnsCorrectResult()
        {
            // Arrange
            int number1 = 12; // 1100
            int number2 = 10; // 1010
            
            // Act
            int result = BitOperations.BitwiseXorInt(number1, number2);
            
            // Assert
            Assert.AreEqual(6, result); // 0110
        }
    }
} 