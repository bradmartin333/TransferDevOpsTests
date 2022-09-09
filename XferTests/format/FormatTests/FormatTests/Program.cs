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
            Console.WriteLine("Press Enter to launch the folder browser or Esc to quit...");

            List<FileInfo> files = new List<FileInfo>();
            while (true)
            {
                ConsoleKey key = Console.ReadKey().Key;
                if (key == ConsoleKey.Enter)
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
                else if (key == ConsoleKey.Escape)
                    break;
            }

            if (files.Count == 0)
            {
                Console.WriteLine("No valid folder selected, exiting...");
                System.Threading.Thread.Sleep(1000);
            }
            else
            {
                FormatExcelDocs(files);
            }

            Console.ReadKey();
        }

        private static void FormatExcelDocs(List<FileInfo> files)
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
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }
        }
    }
}
