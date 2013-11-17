using System;
using System.Collections.Generic;
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

            return table;
        }
    }
}
