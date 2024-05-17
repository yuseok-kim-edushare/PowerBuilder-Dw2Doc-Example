namespace Appeon.DotnetDemo.Dw2Doc.XlsxTester.Exceptions;


[Serializable]
public class AttributeTestException : Exception
{
    public string ObjectName { get; set; }
    public string AttributeName { get; set; }

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
    protected AttributeTestException(
      System.Runtime.Serialization.SerializationInfo info,
      System.Runtime.Serialization.StreamingContext context) : base(info, context)
    {
        ObjectName = string.Empty;
        AttributeName = string.Empty;
    }
}
