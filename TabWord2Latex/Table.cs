using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TabWord2Latex
{
    public enum CellJustification
    {
        Left,
        Center,
        Right
    }
    public enum CellAlignment
    {
        Top,
        Center,
        Bottom
    }

    class Table
    {
        public string Caption { get; set; }
        public bool Hyphenation { get; set; }
        public IEnumerable<Column> Columns { get; set; }
        public Cell[,] Cells { get; set; }

        public int ColsCount
        {
            get
            {
                return Cells == null ? 0 : Cells.GetLength(0);
            }
        }

        public int RowsCount
        {
            get
            {
                return Cells == null ? 0 : Cells.GetLength(1);
            }
        }

        /// <summary>
        /// Calculates row span based on horizontal merge information
        /// and vertical merge based on column span.
        /// </summary>
        internal void CalculateSpan()
        {
            if (Cells == null)
                return;

            // Vertical merge
            for (int i = 0; i < ColsCount; i++)
            {
                Cell mergeCell = null;
                for (int j = 0; j < RowsCount; j++)
                {
                    Cell cell = Cells[i,j];
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

            // Horizontal merge
            for (int i = 0; i < ColsCount; i++)
            {
                for (int j = 0; j < RowsCount; j++)
                {
                    Cell cell = Cells[i, j];
                    if (cell.VMerge == Cell.Merge.Restart &&
                        cell.HMerge == Cell.Merge.Restart)
                    {
                        for (int m = 1; m < cell.ColSpan; m++)
                        {
                            for (int n = 1; n < cell.RowSpan; n++)
                            {
                                Cell mergedCell = Cells[i+m, j+n];
                                mergedCell.VMerge = Cell.Merge.Continue;
                            }
                        }
                    }
                }
            }
        }
    }

    class Column
    {
        /// <summary>
        /// Width in dxa.
        /// </summary>
        public int Width { get; set; }
        //public Column(int dxaWidth)
        //{
        //    Width = dxaWidth / 20; // dxa to pt: http://startbigthinksmall.wordpress.com/2010/01/04/points-inches-and-emus-measuring-units-in-office-open-xml/
        //}
        public Column(int dxaWidth)
        {
            Width = dxaWidth;
        }
    }

    [DebuggerDisplay("{Text}")]
    class Cell
    {
        public enum Merge
        {
            None,
            Restart,
            Continue
        }

        public string Text { get; set; }
        public CellJustification Justification { get; set; }
        public CellAlignment Align { get; set; }

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

        /// <summary>
        /// Width in dxa.
        /// </summary>
        public int Width { get; set; }
    }
}
