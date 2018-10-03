using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PDAImport
{
    public class ItemsShippedLine
    {
        public string itemNo { get; set; }
        public float qtyShipped { get; set; }
        public string uom { get; set; }
    }

    public sealed class ItemsShippedLines
    {
        static List<ItemsShippedLine> _ItemsShippedLines;

        public static List<ItemsShippedLine> ItemsShippedLinesList
        {
            private set { }
            get
            {
                return _ItemsShippedLines;
            }
        }

        static ItemsShippedLines()
        {
            Initialize();
        }

        private static void Initialize()
        {
            _ItemsShippedLines = new List<ItemsShippedLine>();
        }
    }
}
