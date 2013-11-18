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

            // Read caption (previous paragraph to the table)
            var captionPar = wordTable.PreviousSibling<Word.Paragraph>();
            table.Caption = ExtractCaption(captionPar.InnerText);

            //Word.TableProperties tabProp = wordTable.Elements<Word.TableProperties>().First();

            var grid = wordTable.Elements<Word.TableGrid>().First();
            var columns = grid.Elements<Word.GridColumn>();

            table.Columns = columns.Select(gc => new Column(int.Parse(gc.Width.Value)));

            var wordRows = wordTable.Elements<Word.TableRow>();

            int colsTotal = columns.Count();
            int rowsTotal = wordRows.Count();

            var rows = new List<Row>();

            foreach (var wordRow in wordRows)
            {
                var cells = new List<Cell>();
                foreach (var wordCell in wordRow.Elements<Word.TableCell>())
                {
                    var cell = new Cell()
                    {
                        Text = wordCell.InnerText
                    };

                    var cellProp = wordCell.TableCellProperties;
                    if (cellProp != null)
                    {
                        cell.VAlign = ConvertVerticalAlign(cellProp.TableCellVerticalAlignment);

                        if (cellProp.GridSpan != null && cellProp.GridSpan.Val != null
                            && cellProp.GridSpan.Val.HasValue)
                            cell.ColSpan = cellProp.GridSpan.Val.Value;

                        cell.VMerge = ConvertMerge(cellProp.VerticalMerge);

                        cells.Add(cell);

                        // Add additional empty cells for GridSpan value (emulating HorizontalMerge value)
                        for (int i = 1; i < cell.ColSpan; i++)
                            cells.Add(new Cell() { HMerge = Cell.Merge.Continue });
                    }
                }
                rows.Add(new Row { Cells = cells });
            }

            table.Rows = rows;
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
                return Cell.Merge.Continue; // Microsoft Word does not follow specifications
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
