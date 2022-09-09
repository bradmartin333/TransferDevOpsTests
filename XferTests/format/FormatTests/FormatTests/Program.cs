using Microsoft.VisualBasic.FileIO;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace FormatTests
{
    internal class Program
    {
        public static readonly string Project = "Software";

        [STAThread]
        static void Main(string[] args)
        {
            Console.WriteLine("Hello! Are you ready to do some Excel automation?");
            Console.WriteLine("Press Enter to launch the folder browser...");

            List<FileInfo> files = new List<FileInfo>();
            while (true)
            {
                if (Console.ReadKey().Key == ConsoleKey.Enter)
                {
                    FolderBrowserDialog fbd = new FolderBrowserDialog
                    {
                        Description = "Select the downloads folder used for the export TinyTask routine",
                        ShowNewFolderButton = false,
                        RootFolder = Environment.SpecialFolder.MyComputer,
                    };
                    if (fbd.ShowDialog() == DialogResult.OK)
                    {
                        FileInfo[] allFileInfo = Directory.GetFiles(fbd.SelectedPath)
                                                          .Select(x => new FileInfo(x))
                                                          .Where(x => x.Extension.ToLower().Replace(".", "") == "csv")
                                                          .Where(x => x.Name != "output.csv")
                                                          .ToArray();
                        if (allFileInfo.Length > 0) files = allFileInfo.ToList();
                    }
                    break;
                }
                else
                {
                    Console.WriteLine();
                    break;
                }
            }
            
            if (files.Count == 0)
                Console.WriteLine("No valid folder selected");
            else
            {
                Console.WriteLine($"\nI found {files.Count} files!");
                Console.WriteLine("Now just paste in the area path with the right mouse button and press Enter");
                Console.WriteLine("The area path should look something like Software\\InspectRx2\\MedInspSW");
                Console.WriteLine("It can be copied from a Test Case opened in the browser");
                Console.Write("> ");

                string area = Console.ReadLine();

                Console.WriteLine("\nOne moment...");

                while (true)
                {
                    if (FormatExcelDocs(files, area)) break;
                    else
                    {
                        Console.WriteLine("\nHmm I think Excel might be open...");
                        Console.WriteLine("Try closing it then click Enter to try again or any other key to exit...");
                        if (Console.ReadKey().Key == ConsoleKey.Enter) continue;
                        else break;
                    }
                }
                
                Console.WriteLine("All done! Now just import output.csv into the Test Suite in the browser");
            }

            Console.WriteLine("Press any key to exit");
            Console.ReadKey();
        }

        private static bool FormatExcelDocs(List<FileInfo> files, string area)
        {
            try
            {
                string[] firstFileLines = File.ReadAllLines(files[0].FullName);
                string output = firstFileLines[0] + '\n';

                foreach (FileInfo file in files)
                {
                    using (TextFieldParser csvParser = new TextFieldParser(file.FullName))
                    {
                        csvParser.CommentTokens = new string[] { "#" };
                        csvParser.SetDelimiters(new string[] { "," });

                        // Skip the row with the column names
                        csvParser.ReadLine();

                        while (!csvParser.EndOfData)
                        {
                            // Read current line fields, pointer moves to the next line.
                            string[] fields = csvParser.ReadFields();
                            if (!string.IsNullOrEmpty(fields[6])) fields[6] = area;
                            for (int i = 0; i < fields.Length; i++)
                            {
                                if (i == 0 || i == 7)
                                    fields[i] = "";
                                else
                                {
                                    if (fields[i].Contains("\""))
                                        fields[i] = fields[i].Replace("\"", "\"\"");
                                    fields[i] = "\"" + fields[i] + "\"";
                                }
                            }
                            output += string.Join(",", fields) + '\n';
                        }
                    }
                }

                File.WriteAllText(files[0].Directory + @"\output.csv", output);
            }
            catch (Exception)
            {
                return false;
            }
            return true;
        }
    }
}
