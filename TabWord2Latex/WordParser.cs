using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;

namespace TabWord2Latex
{
    class WordParser
    {
        public Table ParseTable(Word.Table wordTable)
        {
            Table table = new Table();
            //Read table caption (first paragraph above the table)
            table.Caption = wordTable.Range.Previous(Word.WdUnits.wdParagraph, 1).Text.TrimEnd(new char[] { '\r', '\n' });
            //Read hyphenation settings
            table.Hyphenation = wordTable.Range.ParagraphFormat.Hyphenation == 1;
            
            //var columns = new List<Column>();

            //foreach (Word.Column wordColumn in wordTable.Columns)
            //    columns.Add(new Column() { Width = wordColumn.Width });

            //table.Columns = columns;

            int totalCols = wordTable.Columns.Count;
            int totalRows = wordTable.Rows.Count;
            for (int i = 1; i <= totalCols; i++)
                for (int j = 1; j <= totalRows; j++)
                {
                    try
                    {
                        Word.Cell cell = wordTable.Cell(j, i);
                        Console.WriteLine("{0} col, {1} row: {2}", i, j, cell.Range.Text);
                    }
                    catch (System.Exception ex)
                    {
                        Console.WriteLine("NO: {0} col, {1} row", i, j);
                    }
                }

            return table;
        }
    }
}
