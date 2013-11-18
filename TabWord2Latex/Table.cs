using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TabWord2Latex
{
    class Table
    {
        public string Caption { get; set; }
        public bool Hyphenation { get; set; }
        public IEnumerable<Column> Columns { get; set; }
        public Cell[,] Cells { get; set; }
    }

    class Column
    {
        public float Width { get; set; }
        public Column(int dxaWidth)
        {
            Width = dxaWidth / 20; //dxa to pt: http://startbigthinksmall.wordpress.com/2010/01/04/points-inches-and-emus-measuring-units-in-office-open-xml/
        }
    }

    class Cell
    {
        public enum HorizontalAlignment
        {
            Left,
            Center,
            Right
        }
        public enum VerticalAlignment
        {
            Top,
            Center,
            Bottom
        }

        public enum Merge
        {
            None,
            Restart,
            Continue
        }

        public string Text { get; set; }
        public HorizontalAlignment HAlign { get; set; }
        public VerticalAlignment VAlign { get; set; }

        public Merge HMerge { get; set; }
        public Merge VMerge { get; set; }

        public int Col { get; set; }
        public int Row { get; set; }
        public int ColSpan { get; set; }
        public int RowSpan { get; set; }
    }
}
