using System;
using System.Collections.Generic;
using System.Globalization;
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

        private static string DxaToPt(int dxa)
        {
            return String.Concat((((float)dxa) / 20).ToString(CultureInfo.InvariantCulture), "pt");
        }

        /// <summary>
        /// Builds column definition for tabular.
        /// </summary>
        private string BuildColDef(Table table)
        {
            StringBuilder s = new StringBuilder("|");
            foreach (var col in table.Columns)
            {
                s.Append("C{").Append(DxaToPt(col.Width)).Append("}|");
            }
            return s.ToString();
        }

        private string BuildJustification(CellJustification just, string value = "")
        {
            string commandName;
            switch (just)
            {
                case CellJustification.Center:
                    commandName = "centering"; break;
                case CellJustification.Right:
                    commandName = "raggedright"; break;
                case CellJustification.Left:
                default:
                    commandName = String.Empty; break;
            }

            return String.IsNullOrEmpty(value) ? Command(commandName) 
                                               : CommandRargs(commandName, value);
        }

        private string BuildTableCells(Table table)
        {
            StringBuilder s = new StringBuilder();
            s.AppendLine(Command("hline"));

            for (int r = 0; r < table.RowsCount; r++)
            {
                bool hasVMergedCells = false;
                for (int c = 0; c < table.ColsCount; c++)
                {
                    Cell cell = table.Cells[c, r];
                    if (cell.HMerge == Cell.Merge.Continue)
                        continue;
                    if (cell.VMerge == Cell.Merge.Continue)
                    {
                        s.Append(" ");
                        hasVMergedCells = true;
                    }
                    else
                    {
                        string value = cell.Text;
                        if (cell.VMerge == Cell.Merge.Restart)
                        {
                            // multirow disables column justification, fix:
                            if (cell.Justification != CellJustification.Left)
                                value = BuildJustification(cell.Justification, value);
                            value = CommandRargs("multirow", cell.RowSpan,
                                Command("hsize"), value);
                            hasVMergedCells = true;
                        }
                        if (cell.HMerge == Cell.Merge.Restart)
                        {
                            value = CommandRargs("multicolumn", cell.ColSpan,
                                "|C{" + DxaToPt(cell.Width) + "}|", value);
                        }

                        if (c != 0)
                            s.Append(" ");
                        s.Append(value).Append(" ");
                    }
                    if (c < table.ColsCount - 1)
                        s.Append("&");
                    
                }

                if (hasVMergedCells)
                {
                    int begin = 0, end = 0;
                    for (int i = 0; i < table.ColsCount; i++)
                    {
                        Cell cell = table.Cells[i, r];
                        //if (cell.HMerge = Cell.Merge.Restart
                        s.Append(CommandRargs("tncl"));
                    }
                }
                else
                {
                    s.Append(Command("tnhl"));
                }
                
                if (r < table.RowsCount - 1)
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
                s.AppendLine(BuildTableCells(table));
                s.AppendLine(EndEnvironment("tabular"));
            }
            // \end{tabular}

            s.AppendLine(EndEnvironment("table"));
            return s.ToString();
        }
    }
}
