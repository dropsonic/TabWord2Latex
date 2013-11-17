using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;

namespace TabWord2Latex
{
    class Converter
    {
        /// <summary>
        /// Builds a command with required args only.
        /// </summary>
        public static string CommandRargs(string name, params object[] args)
        {
            StringBuilder s = new StringBuilder();
            s.Append(@"\").Append(name);
            foreach(var arg in args)
                s.Append(@"{").Append(arg).Append(@"}");
            return s.ToString();
        }

        public static string Command(string name)
        {
            return String.Concat(@"\", name);
        }

        public static string BeginEnvironment(string envname, params object[] args)
        {
            StringBuilder s = new StringBuilder();
            s.Append(@"\begin{").Append(envname).Append(@"}");
            foreach (var arg in args)
                s.Append(@"{").Append(arg).Append(@"}");
            return s.ToString();
        }

        public static string EndEnvironment(string envname)
        {
            return CommandRargs("end", envname);
        }

        public static string ToTex(Word.Table table)
        {
            StringBuilder s = new StringBuilder();
            //Read table caption (first paragraph above the table)
            string caption = table.Range.Previous(Word.WdUnits.wdParagraph, 1).Text.TrimEnd(new char[] {'\r', '\n'});
            //Read hyphenation settings
            bool hyphenation = table.Range.ParagraphFormat.Hyphenation == 1;

            //float leftPadding = table.LeftPadding;
            //float rightPadding = table.RightPadding;
            //var lpadCmd = CommandRargs("hspace", leftPadding);
            //var rpadCmd = CommandRargs("hspace", rightPadding);

            s.AppendLine(BeginEnvironment("table"));
                s.AppendLine(CommandRargs("caption", caption));
                s.AppendLine(CommandRargs("label", "tab:labelname"));
                if (!hyphenation)
                    s.AppendLine(Command("nohyphenation"));
                s.AppendLine();
                

            s.AppendLine(EndEnvironment("table"));
            return s.ToString();
        }
    }
}
