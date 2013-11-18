using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Word = DocumentFormat.OpenXml.Wordprocessing;

namespace TabWord2Latex
{
    class WordParser
    {
        public Table ParseTable(Word.Table wordTable)
        {
            Table table = new Table();

            //Read caption (previous paragraph to the table)
            var captionPar = wordTable.PreviousSibling<Word.Paragraph>();
            table.Caption = ExtractCaption(captionPar.InnerText);

            //Word.TableProperties tabProp = wordTable.Elements<Word.TableProperties>().First();

            var grid = wordTable.Elements<Word.TableGrid>().First();
            var columns = grid.Elements<Word.GridColumn>();

            table.Columns = columns.Select(gc => new Column(int.Parse(gc.Width.Value)));

            var rows = wordTable.Elements<Word.TableRow>();

            int colsTotal = columns.Count();
            int rowsTotal = rows.Count();
            table.Cells = new Cell[colsTotal, rowsTotal];

            int r = 0;
            foreach (var row in rows)
            {
                int c = 0;

                foreach (var wordCell in row.Elements<Word.TableCell>())
                {
                    var cell = new Cell()
                    {
                        Text = wordCell.InnerText,
                        Col = c,
                        Row = r
                    };
                    //Console.WriteLine(cell.InnerText);
                    var cellProp = wordCell.Elements<Word.TableCellProperties>().First();

                    cell.VAlign = ConvertVerticalAlign(cellProp.TableCellVerticalAlignment);

                    cell.HMerge = ConvertMerge(cellProp.HorizontalMerge);
                    cell.VMerge = ConvertMerge(cellProp.VerticalMerge);

                    table.Cells[c, r] = cell;
                    c++;
                }
                r++;
            }

            table.CalculateSpan();
            return table;
        }

        /// <summary>
        /// Extracts table name from full caption (excludes "table #N" part).
        /// </summary>
        static string ExtractCaption(string captionString)
        {
            Regex regex = new Regex(@"^(Таблица|Table)\s*[0-9,.:;-]*\s*[-–—]*\s*", RegexOptions.Singleline | RegexOptions.IgnoreCase);
            var match = regex.Match(captionString);
            if (match.Success)
                return captionString.Remove(match.Index, match.Length);
            else
                return "";
        }

        static Cell.VerticalAlignment ConvertVerticalAlign(Word.TableCellVerticalAlignment align)
        {
            if (align == null)
                return Cell.VerticalAlignment.Center; //default value
            switch (align.Val.Value)
            {
                case Word.TableVerticalAlignmentValues.Bottom:
                    return Cell.VerticalAlignment.Bottom;
                case Word.TableVerticalAlignmentValues.Center:
                    return Cell.VerticalAlignment.Center;
                case Word.TableVerticalAlignmentValues.Top:
                    return Cell.VerticalAlignment.Top;
                default:
                    return Cell.VerticalAlignment.Center;
            }
        }

        static Cell.Merge ConvertMergeValue(Word.MergedCellValues value)
        {
            switch (value)
            {
                case Word.MergedCellValues.Restart:
                    return Cell.Merge.Restart;
                case Word.MergedCellValues.Continue:
                    return Cell.Merge.Continue;
                default:
                    return Cell.Merge.None;
            }
        }

        static Cell.Merge ConvertMerge(Word.VerticalMerge merge)
        {
            if (merge == null)
                return Cell.Merge.None;
            else if (merge.Val == null)
                return Cell.Merge.Continue; //Microsoft Word does not follow specifications
            else
                return ConvertMergeValue(merge.Val.Value);
        }

        static Cell.Merge ConvertMerge(Word.HorizontalMerge merge)
        {
            if (merge == null)
                return Cell.Merge.None;
            else if (merge.Val == null)
                return Cell.Merge.Continue; //Microsoft Word does not follow specifications
            else
                return ConvertMergeValue(merge.Val.Value);
        }

        /// <summary>
        /// Returns width in dxa (not pt!).
        /// </summary>
        /// <returns>Width in dxa.</returns>
        static int ConvertWidth(Word.TableCellWidth width)
        {
            if (width == null || width.Width == null)
                return 0;
            else
                return int.Parse(width.Width);
        }
    }
}
