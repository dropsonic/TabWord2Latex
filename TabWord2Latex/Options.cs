using CommandLine;
using CommandLine.Text;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TabWord2Latex
{
    class Options
    {
        [ParserState]
        public IParserState LastParserState { get; set; }

        [Option('i', "input", Required = true,
            HelpText = "Word document with tables to be processed.")]
        [ValueOption(0)]
        public string WordFileName { get; set; }

        [Option('o', "output", Required = true,
            HelpText = "TeX output file for converted table.")]
        [ValueOption(1)]
        public string TexFileName { get; set; }

        [Option('t', "table", Required = false, DefaultValue = 1,
            HelpText = "Number of table to be converted.")]
        [ValueOption(2)]
        public int TableNumber { get; set; }

        [HelpOption('h', "help")]
        public string GetUsage()
        {
            var help = new HelpText
            {
                Heading = new HeadingInfo("TabWord2Latex", "1.0"),
                Copyright = new CopyrightInfo("Vladimir Panchenko", 2013),
                AdditionalNewLineAfterOption = true,
                AddDashesToOption = true
            };
            help.AddPreOptionsLine("Usage: tabword2latex -i <input word file> -o <output tex file> -t [table number]");
            help.AddOptions(this);

            if (LastParserState.Errors.Count > 0)
            {
                var errors = help.RenderParsingErrorsText(this, 2); // indent with two spaces
                if (!string.IsNullOrEmpty(errors))
                {
                    help.AddPreOptionsLine(string.Concat(Environment.NewLine, "ERROR(S):"));
                    help.AddPreOptionsLine(errors);
                }
            }

            return help;
        }
    }
}
