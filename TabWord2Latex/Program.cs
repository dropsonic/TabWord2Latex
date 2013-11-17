using CommandLine;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;

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
                Word.Document document = null;
                Word.Application app = null;
                try
                {
                    app = new Word.Application();
                    document = app.Documents.Open(options.WordFileName);

                    if (document.Tables.Count <= 0)
                        throw new ApplicationException("Document does not contain a table.");

                    if (options.TableNumber > document.Tables.Count)
                        throw new ApplicationException(String.Format("Table number is too big. Document contains only {0} tables.", document.Tables.Count));

                    Word.Table table = document.Tables[options.TableNumber];
                    //foreach (Word.Paragraph item in document.Paragraphs)
                    //    Console.WriteLine(item.Range.Text);
                    Debug.WriteLine(table.LeftPadding);
                    Debug.WriteLine(table.RightPadding);
                    Debug.WriteLine(table.TopPadding);
                    Debug.WriteLine(table.BottomPadding);

                    string texTable = Converter.ToTex(table);
                    File.WriteAllText(options.TexFileName, texTable, Encoding.UTF8);

                    Console.WriteLine("Done.");
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
                    if (document != null)
                        document.Close();
                    if (app != null)
                        app.Quit();
                }
            }
        }
    }
}
