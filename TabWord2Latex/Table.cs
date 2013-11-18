using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TabWord2Latex
{
    class Cell
    {
        public float Width { get; set; }
        public float Height { get; set; }
        public string Text { get; set; }

        public int Col { get; set; }
        public int Row { get; set; }
        public int ColSpan { get; set; }
        public int RowSpan { get; set; }
    }

    class Table
    {
        public string Caption { get; set; }

        public bool Hyphenation { get; set; }

        //public IEnumerable<Column> Columns { get; set; }
        public Cell[,] Cells { get; set; }
    }
}
