using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BatchPowerPointToPDF
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                Console.Write("Type full path (including filename) of PowerPoint file: ");

                string powerPointFilename = Console.ReadLine();

                Console.Write("Path and filename of exported PDF: ");

                string outputFilename = Console.ReadLine();

                PowerPointToPDFExporter.ConvertToPDF(powerPointFilename, outputFilename);

                Console.WriteLine("Successfully converted!");
                Console.ReadLine();
            }

            catch (Exception e)
            {
                Console.WriteLine("Critical Error: " + e.Message);
                Console.ReadLine();
            }
        }
    }
}
