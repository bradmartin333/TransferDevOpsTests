using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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
            while (true)
            {
                ConsoleKey key = Console.ReadKey().Key;
                if (key == ConsoleKey.Enter)
                {
                    continue;
                }
                else if (key == ConsoleKey.Escape)
                {
                    break;
                }
            }
        }
    }
}
