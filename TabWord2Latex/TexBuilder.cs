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

        /// <summary>
        /// Builds column definition for tabular.
        /// </summary>
        private string BuildColDef(Table table)
        {
            StringBuilder s = new StringBuilder("|");
            foreach (var col in table.Columns)
            {
                s.Append("C{").Append(col.Width).Append("pt}|");
            }
            return s.ToString();
        }

        private string BuildTableRow(Row row)
        {
            StringBuilder s = new StringBuilder("|");
            foreach (var cell in row.Cells)
            {
                if (cell.HMerge == Cell.Merge.Continue ||
                    cell.VMerge == Cell.Merge.Continue)
                {
                    s.Append(" &");
                }
                else
                {
                    //if (cell.
                }
                s.AppendLine();
            }
            return s.ToString();
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

            // \begin{tabular}{...}
            {
                s.AppendLine(BeginEnvironment("tabular", BuildColDef(table)));
                s.AppendLine(CommandRargs("hline"));
                foreach (var row in table.Rows)
                    s.AppendLine(BuildTableRow(row));
                s.AppendLine(EndEnvironment("tabular"));
            }
            // \end{tabular}

            s.AppendLine(EndEnvironment("table"));
            return s.ToString();
        }
    }
}
