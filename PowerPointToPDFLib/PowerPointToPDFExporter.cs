using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace PowerPointToPDFLib
{
    public class PowerPointToPDFExporter
    {
        public static void ConvertToPDF(string pptFilename, string outputFilename)
        {
            try
            {
                // Sanitizing filename strings
                pptFilename = pptFilename.Replace(@"\\", @"\");
                outputFilename = outputFilename.Replace(@"\\", @"\");

                PowerPoint.Application app = new PowerPoint.Application();
                var presentation = app.Presentations;

                // Opening PowerPoint
                var file = app.Presentations.Open(pptFilename, MsoTriState.msoTrue, MsoTriState.msoTrue, MsoTriState.msoFalse);

                // Converting to PDF
                file.ExportAsFixedFormat(outputFilename, PowerPoint.PpFixedFormatType.ppFixedFormatTypePDF);
            }

            catch (Exception e)
            {
                // TODO: Implement better exception handling.
                Console.WriteLine("Critical Failure: " + e.Message);
            }

            finally
            {
                // TODO: Perform any possible cleanup after failure.
            }
        }
    }
}
