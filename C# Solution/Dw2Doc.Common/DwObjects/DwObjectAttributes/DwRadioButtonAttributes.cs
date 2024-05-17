namespace Appeon.DotnetDemo.Dw2Doc.Common.DwObjects.DwObjectAttributes;

public class DwRadioButtonAttributes : DwTextAttributes
{
    public IDictionary<string, string>? CodeTable { get; set; }
    public short Columns { get; set; }
    public bool LeftText { get; set; }

    // override object.Equals
    public override bool Equals(object obj)
    {
        return obj is DwRadioButtonAttributes other
            && Columns == other.Columns
            && LeftText == other.LeftText
            && ((CodeTable is null) == (other.CodeTable is null))
            && (CodeTable is null || CodeTable.Keys.SequenceEqual(other.CodeTable!.Keys))
            && (CodeTable is null || CodeTable.Values.SequenceEqual(other.CodeTable!.Values));

    }

}
