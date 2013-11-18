using System;
using System.Collections.Generic;
using System.Diagnostics;
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
        public int ColsCount
        {
            get { return Cells.GetLength(0); }
        }
        public int RowsCount
        {
            get { return Cells.GetLength(1); }
        }

        /// <summary>
        /// Calculates span based on merge information.
        /// </summary>
        internal void CalculateSpan()
        {
            if (Cells == null)
                return;

            //Vertical merge
            for (int i = 0; i < ColsCount; i++)
            {
                Cell mergeCell = null;
                for (int j = 0; j < RowsCount; j++)
                {
                    Cell cell = Cells[i, j];
                    cell.RowSpan = 1;
                    switch (cell.VMerge)
                    {
                        case Cell.Merge.Restart:
                            mergeCell = cell; 
                            break;
                        case Cell.Merge.Continue:
                            if (mergeCell != null)
                                mergeCell.RowSpan++;
                            break;
                    }
                }
            }

            //Horizontal merge
            for (int j = 0; j < RowsCount; j++)
            {
                Cell mergeCell = null;
                for (int i = 0; i < ColsCount; i++)
                {
                    Cell cell = Cells[i, j];
                    cell.ColSpan = 1;
                    switch (cell.HMerge)
                    {
                        case Cell.Merge.Restart:
                            mergeCell = cell;
                            break;
                        case Cell.Merge.Continue:
                            if (mergeCell != null)
                                mergeCell.ColSpan++;
                            break;
                    }
                }
            }
        }
    }

    class Column
    {
        public float Width { get; set; }
        public Column(int dxaWidth)
        {
            Width = dxaWidth / 20; //dxa to pt: http://startbigthinksmall.wordpress.com/2010/01/04/points-inches-and-emus-measuring-units-in-office-open-xml/
        }
    }

    [DebuggerDisplay("{Text}")]
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

        private int _colSpan = 1;
        public int ColSpan
	    {
		    get { return _colSpan;}
		    set { _colSpan = value;}
	    }
	
        private int _rowSpan = 1;
        public int RowSpan
	    {
		    get { return _rowSpan;}
		    set { _rowSpan = value;}
	    }
    }
}
