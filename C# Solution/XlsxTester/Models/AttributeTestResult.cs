namespace Appeon.DotnetDemo.Dw2Doc.XlsxTester.Models;

public class AttributeTestResult
{
    public string ObjectName { get; set; }
    public string Attribute { get; set; }
    public bool Result => ExpectedValue == RealValue
        || (decimal.TryParse(ExpectedValue, out decimal dec1)
            && decimal.TryParse(RealValue, out decimal dec2) && dec1 == dec2);
    public string? ExpectedValue { get; set; }
    public string? RealValue { get; set; }

    public AttributeTestResult(string objectName, string attribute, string? expectedValue, string? realValue)
    {
        ObjectName = objectName;
        Attribute = attribute;
        ExpectedValue = expectedValue;
        RealValue = realValue;
    }
}
