using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;

namespace TabWord2Latex
{
    public static class Converter
    {
        public static string ToTex(Word.Table wordTable)
        {
            var parser = new WordParser();
            var builder = new TexBuilder();
            var table = parser.ParseTable(wordTable);
            return builder.TableToTex(table);
        }
    }
}
