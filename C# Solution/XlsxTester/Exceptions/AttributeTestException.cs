using System.Diagnostics.CodeAnalysis;
using System.Runtime.Serialization;

namespace Appeon.DotnetDemo.Dw2Doc.XlsxTester.Exceptions;

[Serializable]
public class AttributeTestException : Exception, ISerializable
{
    public string ObjectName { get; set; } = string.Empty;
    public string AttributeName { get; set; } = string.Empty;

    public AttributeTestException(string objectName, string propName) : base()
    {
        ObjectName = objectName;
        AttributeName = propName;
    }

    public AttributeTestException(string message, string objectName, string propName) : base(message)
    {
        ObjectName = objectName;
        AttributeName = propName;
    }
    
    public AttributeTestException(
        string objectName, string propName,
        string message, Exception inner) : base(message, inner)
    {
        ObjectName = objectName;
        AttributeName = propName;
    }
    
#if NET8_0_OR_GREATER
    [Obsolete("This API supports obsolete formatter-based serialization. It should not be called or extended by application code.", DiagnosticId = "SYSLIB0051")]
#endif
    protected AttributeTestException(SerializationInfo info, StreamingContext context)
        : base(info, context)
    {
        ObjectName = info.GetString(nameof(ObjectName)) ?? string.Empty;
        AttributeName = info.GetString(nameof(AttributeName)) ?? string.Empty;
    }
    
#if NET8_0_OR_GREATER
    [Obsolete("This API supports obsolete formatter-based serialization. It should not be called or extended by application code.", DiagnosticId = "SYSLIB0051")]
#endif
    void ISerializable.GetObjectData(SerializationInfo info, StreamingContext context)
    {
        base.GetObjectData(info, context);
        info.AddValue(nameof(ObjectName), ObjectName);
        info.AddValue(nameof(AttributeName), AttributeName);
    }
}
