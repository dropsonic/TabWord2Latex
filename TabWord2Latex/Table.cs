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
        public IList<Row> Rows { get; set; }

        /// <summary>
        /// Calculates span based on merge information.
        /// </summary>
        internal void CalculateSpan()
        {
            if (Rows == null)
                return;

            int colsCount = Columns.Count();
            // Vertical merge
            for (int i = 0; i < colsCount; i++)
            {
                Cell mergeCell = null;
                for (int j = 0; j < Rows.Count; j++)
                {
                    Cell cell = Rows[j].Cells[i];
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
        }
    }

    class Row
    {
        public IList<Cell> Cells { get; set; }
        ///// <summary>
        ///// Alignment for most of the cells.
        ///// </summary>
        //public CellAlignment RowAlignment
        //{
        //    get
        //    {
        //        // not optimal, but easy to read
        //        int top = Cells.Count(c => c.Align == CellAlignment.Top);
        //        int center = Cells.Count(c => c.Align == CellAlignment.Center);
        //        int bottom = Cells.Count(c => c.Align == CellAlignment.Bottom);
        //        CellAlignment result = CellAlignment.Center;
        //        if (top > center)
        //            if (top > bottom)
        //                result = CellAlignment.Top;
        //            else
        //                result = CellAlignment.Bottom;
        //        else
        //            if (bottom > center)
        //                result = CellAlignment.Bottom;
        //            else
        //                result = CellAlignment.Center;
        //        return result;
        //    }
        //}
        ///// <summary>
        ///// Justification for most of the cells.
        ///// </summary>
        //public CellJustification RowJustification
        //{
        //    get
        //    {
        //        int left = Cells.Count(c => c.Justification == CellJustification.Left);
        //        int center = Cells.Count(c => c.Justification == CellJustification.Center);
        //        int bottom = Cells.Count(c => c.Justification == CellJustification.Right);
        //        CellJustification result = CellJustification.Center;
        //        if (left > center)
        //            if (left > bottom)
        //                result = CellJustification.Left;
        //            else
        //                result = CellJustification.Right;
        //        else
        //            if (bottom > center)
        //                result = CellJustification.Right;
        //            else
        //                result = CellJustification.Center;
        //        return result;
        //    }
        //}
    }

    class Column
    {
        public float Width { get; set; }
        public Column(int dxaWidth)
        {
            Width = dxaWidth / 20; // dxa to pt: http://startbigthinksmall.wordpress.com/2010/01/04/points-inches-and-emus-measuring-units-in-office-open-xml/
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
    }
}
