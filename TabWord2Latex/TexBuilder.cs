using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TabWord2Latex
{
    class TexBuilder
    {
        /// <summary>
        /// Builds a command with required args only.
        /// </summary>
        private string CommandRargs(string name, params object[] args)
        {
            StringBuilder s = new StringBuilder();
            s.Append(@"\").Append(name);
            foreach (var arg in args)
                s.Append(@"{").Append(arg).Append(@"}");
            return s.ToString();
        }

        private string Command(string name)
        {
            return String.Concat(@"\", name);
        }

        private string BeginEnvironment(string envname, params object[] args)
        {
            StringBuilder s = new StringBuilder();
            s.Append(@"\begin{").Append(envname).Append(@"}");
            foreach (var arg in args)
                s.Append(@"{").Append(arg).Append(@"}");
            return s.ToString();
        }

        private string EndEnvironment(string envname)
        {
            return CommandRargs("end", envname);
        }

        public string TableToTex(Table table)
        {
            StringBuilder s = new StringBuilder();

            //float leftPadding = table.LeftPadding;
            //float rightPadding = table.RightPadding;
            //var lpadCmd = CommandRargs("hspace", leftPadding);
            //var rpadCmd = CommandRargs("hspace", rightPadding);

            s.AppendLine(BeginEnvironment("table"));
            s.AppendLine(CommandRargs("caption", table.Caption));
            s.AppendLine(CommandRargs("label", "tab:labelname"));
            if (table.Hyphenation == false)
                s.AppendLine(Command("nohyphenation"));
            s.AppendLine();


            s.AppendLine(EndEnvironment("table"));
            return s.ToString();
        }
    }
}
