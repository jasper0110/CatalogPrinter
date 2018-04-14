using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CatalogPrinter
{
    public class CatalogType
    {
        public string Text { get; }
        public int Value { get; }

        public CatalogType(string t, int v)
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
