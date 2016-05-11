using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
//using System.Threading. Tasks;

namespace PrintWindowsService
{
    class Program
    {
        static void Main(string[] args)
        {
            PrintJobs pJobs = new PrintJobs();
            pJobs.StartJob();
            Console.WriteLine("Press Esc to exit");
            do
            {
            } while (Console.ReadKey().Key != ConsoleKey.Escape);
            pJobs.StopJob();
        }
    }
}
