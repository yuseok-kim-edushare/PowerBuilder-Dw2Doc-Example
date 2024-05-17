﻿using System.Drawing;

namespace Appeon.DotnetDemo.Dw2Doc.Common.DwObjects
{
    /// <summary>
    /// Simple wrapper for a <see cref="Color"/>
    /// </summary>
    public class DwColorWrapper
    {
        public Color Value { get; set; }

        // override object.Equals
        public override bool Equals(object obj)
        {
            return obj is DwColorWrapper other
                && other.Value == Value;
        }

        // override object.GetHashCode
        public override int GetHashCode()
        {
            return Value.GetHashCode();
        }
    }
}
