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
        /// Returns column definition with specific align, justification and width.
        /// </summary>
        /// <param name="width">Width in dxa.</param>
        private string GetColDef(CellAlignment align, CellJustification justfn, int width)
        {
            // TODO: add alignment and justification support
            // (strategy pattern or align+justfn struct and pattern matching through dictionary)
            StringBuilder s = new StringBuilder();
            s.Append("C{").Append(DxaToPt(width)).Append("}");
            return s.ToString();
        }

        /// <summary>
        /// Builds column definition for tabular.
        /// </summary>
        private string BuildColDef(Table table)
        {
            StringBuilder s = new StringBuilder("|");
            foreach (var col in table.Columns)
            {
                // TODO: add alignment and justification support
                s.Append(GetColDef(CellAlignment.Center, CellJustification.Center, col.Width)).Append("|");
            }
            return s.ToString();
        }

        private string BuildJustification(CellJustification justfn, string value = "")
        {
            string commandName;
            switch (justfn)
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

            for (int r = 0; r < table.RowsCount; r++)
            {
                // Line length is calculated for the top line of the row
                int begin = 0, end = 0;
                for (int i = 0; i < table.ColsCount; i++)
                {
                    Cell cell = table.Cells[i, r];
                    if (cell.VMerge == Cell.Merge.Continue)
                    {
                        begin++;
                        end++;
                        if (end - 1 - begin > 0) // if we are at solid row now
                        {
                            s.Append(CommandRargs("cline", String.Format("{0}-{1}", begin, end-1)));
                            begin = end-1;
                        }
                    }
                    else
                    {
                        end++;
                    }
                }
                if (begin == 0 && end == table.ColsCount)
                    s.Append(Command("hline"));
                s.AppendLine();
                    
                for (int c = 0; c < table.ColsCount; c++)
                {
                    Cell cell = table.Cells[c, r];

                    if (cell.HMerge == Cell.Merge.Continue)
                    {
                        continue;
                    }
                    if (cell.VMerge == Cell.Merge.Continue)
                    {
                        if (cell.HMerge == Cell.Merge.Restart)
                            s.Append(CommandRargs("multicolumn", cell.ColSpan, (c == 0 ? "|" : "") +
                                GetColDef(cell.Align, cell.Justification, cell.Width) + "|", String.Empty))
                            .Append(" ");
                        else
                            s.Append(" ");
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
                        }
                        if (cell.HMerge == Cell.Merge.Restart)
                        {
                            value = CommandRargs("multicolumn", cell.ColSpan, (c == 0 ? "|" : "") +
                                GetColDef(cell.Align, cell.Justification, cell.Width) + "|", value);
                        }

                        if (c != 0)
                            s.Append(" ");
                        s.Append(value).Append(" ");
                    }

                    s.Append(c < (table.ColsCount - 1) ? "&" : @"\\");
                }
                
                if (r < table.RowsCount - 1)
                    s.AppendLine();
            }

            s.AppendLine();
            s.Append(Command("hline"));
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
