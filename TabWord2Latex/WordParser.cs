﻿using System;
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

            table.Columns = columns.Select(gc => new Column(int.Parse(gc.Width.Value))).ToList();

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
                    var par = wordCell.Descendants<Word.Paragraph>().First();
                    cell.Justification = ParseJustification(par);

                    // Parsing cell properties
                    var cellProp = wordCell.TableCellProperties;
                    if (cellProp != null)
                    {
                        cell.Align = ConvertAlign(cellProp.TableCellVerticalAlignment);

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

        static CellAlignment ConvertAlign(Word.TableCellVerticalAlignment align)
        {
            if (align != null && align.Val != null && align.Val.HasValue)
            {
                switch (align.Val.Value)
                {
                    case Word.TableVerticalAlignmentValues.Bottom:
                        return CellAlignment.Bottom;
                    case Word.TableVerticalAlignmentValues.Top:
                        return CellAlignment.Top;
                    case Word.TableVerticalAlignmentValues.Center:
                    default:
                        return CellAlignment.Center;
                }
            }

            return CellAlignment.Center; //default value
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
            else if (merge.Val == null || !merge.Val.HasValue)
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

        static CellJustification ParseJustification(Word.Paragraph par)
        {
            if (par != null && par.ParagraphProperties != null &&
                par.ParagraphProperties.Justification != null &&
                par.ParagraphProperties.Justification.Val != null &&
                par.ParagraphProperties.Justification.Val.HasValue)
            {
                switch (par.ParagraphProperties.Justification.Val.Value)
                {
                    case Word.JustificationValues.Left:
                        return CellJustification.Left;
                    case Word.JustificationValues.Right:
                        return CellJustification.Right;
                    case Word.JustificationValues.Center:
                    default:
                        return CellJustification.Center;
                }
            }

            return CellJustification.Left; // default value for Microsoft Word
        }
    }
}
