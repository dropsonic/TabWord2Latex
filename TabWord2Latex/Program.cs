using CommandLine;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Word = DocumentFormat.OpenXml.Wordprocessing;

namespace TabWord2Latex
{
    class Program
    {
        static void Main(string[] args)
        {
            var options = new Options();
            if (CommandLine.Parser.Default.ParseArguments(args, options))
            {
                options.WordFileName = Path.GetFullPath(options.WordFileName);

                if (!File.Exists(options.WordFileName))
                {
                    Console.WriteLine("File not found ({0}).", options.WordFileName);
                    return;
                }
                try
                {
                    using (WordprocessingDocument document = WordprocessingDocument.Open(options.WordFileName, false))
                    {
                        var body = document.MainDocumentPart.Document.Body;

                        var tables = body.Elements<Word.Table>();
                        var tablesCount = tables.Count();

                        if (tablesCount <= 0)
                            throw new ApplicationException("Document does not contain a table.");

                        if (options.TableNumber > tablesCount)
                            throw new ApplicationException(String.Format("Table number is too big. Document contains only {0} tables.", tablesCount));

                        var table = tables.ElementAt(options.TableNumber - 1);

                        string texTable = Converter.ToTex(table);
                        File.WriteAllText(options.TexFileName, texTable, Encoding.UTF8);

                        Console.WriteLine("Done.");
                    }
                }
                catch (ApplicationException aex)
                {
                    Console.WriteLine(aex.Message);
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Error: {0} - {1}", ex.GetType().Name, ex.Message);
                }
                finally
                {
                }
            }
        }
    }
}
