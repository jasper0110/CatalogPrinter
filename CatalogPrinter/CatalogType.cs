using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CatalogPrinter
{
    public enum CatalogTypeEnum
    {
        NONSELECTED,
        DAKWERKER,
        VERANDA,
        AANNEMER,
        PARTICULIER
    }

    public class CatalogType
    {
        public string Text { get; }
        public CatalogTypeEnum Value { get; }

        public CatalogType(string t, CatalogTypeEnum v)
        {
            Text = t;
            Value = v;
        }

        public override string ToString()
        {
            return Text;
        }
    }
}
