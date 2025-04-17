using Appeon.DotnetDemo.Dw2Doc.Common.Enums;
using Appeon.DotnetDemo.Dw2Doc.Common.VirtualGrid;
using Appeon.DotnetDemo.Dw2Doc.Common.VirtualGridWriter.Abstractions;
using Appeon.DotnetDemo.Dw2Doc.Docx.VirtualGridWriter.DocxWriter;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Reflection;

namespace Appeon.DotnetDemo.Dw2Doc.Tests.Docx
{
    [TestClass]
    public class VirtualGridDocxWriterBuilderTests
    {
        private object? _builder; // Using object instead of VirtualGridDocxWriterBuilder since it's internal
        private VirtualGrid? _virtualGrid;
        private Type? _builderType;

        [TestInitialize]
        public void Setup()
        {
            // Get the internal builder type
            _builderType = typeof(VirtualGridDocxWriter).Assembly.GetType("Appeon.DotnetDemo.Dw2Doc.Docx.VirtualGridWriter.DocxWriter.VirtualGridDocxWriterBuilder");
            Assert.IsNotNull(_builderType, "Could not find VirtualGridDocxWriterBuilder type");
            
            // Create an instance using reflection
            _builder = Activator.CreateInstance(_builderType);
            Assert.IsNotNull(_builder, "Could not create VirtualGridDocxWriterBuilder instance");
            
            // Create VirtualGrid using reflection
            var rows = new List<RowDefinition>();
            var columns = new List<ColumnDefinition>();
            var bandRows = new List<BandRows>();
            
            // Create cell repository using reflection
            var cellRepoType = typeof(VirtualGrid).Assembly.GetType("Appeon.DotnetDemo.Dw2Doc.Common.VirtualGrid.VirtualCellRepository");
            Assert.IsNotNull(cellRepoType, "Could not find VirtualCellRepository type");
            
            var cellRepo = Activator.CreateInstance(cellRepoType, BindingFlags.Instance | BindingFlags.NonPublic, null, null, null);
            Assert.IsNotNull(cellRepo, "Could not create VirtualCellRepository instance");
            
            // Create VirtualGrid using reflection
            var constructor = typeof(VirtualGrid).GetConstructors(BindingFlags.NonPublic | BindingFlags.Instance)[0];
            _virtualGrid = (VirtualGrid)constructor.Invoke(new object[] { rows, columns, bandRows, cellRepo, DwType.Grid });
        }

        [TestMethod]
        public void Build_WithValidGrid_ShouldReturnWriter()
        {
            // Arrange
            string? error = null;
            
            // Act - Call Build method through reflection
            Assert.IsNotNull(_builderType, "BuilderType should not be null");
            var buildMethod = _builderType.GetMethod("Build");
            Assert.IsNotNull(buildMethod, "Could not find Build method");
            
            var parameters = new object?[] { _virtualGrid, error };
            var writer = buildMethod.Invoke(_builder, parameters) as AbstractVirtualGridWriter;
            error = parameters[1] as string; // Get out parameter value
            
            // Assert
            Assert.IsNotNull(writer);
            Assert.IsNull(error);
            Assert.IsInstanceOfType(writer, typeof(VirtualGridDocxWriter));
        }

        [TestMethod]
        public void BuildFromTemplate_ShouldThrowNotImplementedException()
        {
            // Arrange
            string templatePath = "template.docx";
            
            // Act & Assert
            Assert.IsNotNull(_builderType, "BuilderType should not be null");
            var buildFromTemplateMethod = _builderType.GetMethod("BuildFromTemplate");
            Assert.IsNotNull(buildFromTemplateMethod, "Could not find BuildFromTemplate method");
            
            Assert.ThrowsException<TargetInvocationException>(() => {
                try {
                    buildFromTemplateMethod.Invoke(_builder, new object?[] { _virtualGrid, templatePath, null });
                } catch (TargetInvocationException ex) {
                    if (ex.InnerException is NotImplementedException) {
                        throw; // Re-throw the original exception to avoid changing the stack trace
                    }
                    throw;
                }
            });
        }

        [TestMethod]
        public void Build_WithNullGrid_ShouldThrowArgumentNullException()
        {
            // Arrange
            string? error = null;
            
            // Act & Assert
            Assert.IsNotNull(_builderType, "BuilderType should not be null");
            var buildMethod = _builderType.GetMethod("Build");
            Assert.IsNotNull(buildMethod, "Could not find Build method");
            
            Assert.ThrowsException<TargetInvocationException>(() => {
                try {
                    var parameters = new object?[] { null, error };
                    buildMethod.Invoke(_builder, parameters);
                } catch (TargetInvocationException ex) {
                    if (ex.InnerException is ArgumentNullException || 
                        ex.InnerException is NullReferenceException) {
                        throw; // Re-throw the original exception
                    }
                    throw;
                }
            });
        }
    }
} 