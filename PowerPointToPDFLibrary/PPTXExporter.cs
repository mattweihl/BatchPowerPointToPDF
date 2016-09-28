using System;
using Microsoft.Office.Core;
using Microsoft.Win32;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointToPDFLibrary
{
    public class PptxExporter
    {
        /// <summary>
        /// Converts the given PowerPoint presentation to a PDF.
        /// </summary>
        /// <param name="pptFilename"></param>
        public static void ConvertToPdf(string pptFilename)
        {
            try
            {
                if (!OfficeInstalled())
                {
                    throw new NotSupportedException("Office Installation Not Detected");
                }

                // Adding Escape Characters
                pptFilename = pptFilename.Replace(@"\\", @"\");

                PowerPoint.Application app = new PowerPoint.Application();
                PowerPoint.Presentations presentation = app.Presentations;

                // Opening PowerPoint
                PowerPoint.Presentation file = app.Presentations.Open(pptFilename, MsoTriState.msoTrue, MsoTriState.msoTrue, MsoTriState.msoFalse);

                // Converting to PDF
                file.ExportAsFixedFormat(pptFilename + ".pdf", PowerPoint.PpFixedFormatType.ppFixedFormatTypePDF);
            }

            catch (Exception e)
            {
                // TODO: Implement better exception handling.
                Console.WriteLine("Critical Failure: " + e.Message);
            }
        }

        /// <summary>
        /// Returns whether Office is installed on the Local Machine.
        /// </summary>
        /// <returns></returns>
        public static bool OfficeInstalled()
        {
            RegistryKey powerpointKey = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\powerpnt.exe");

            powerpointKey?.Close();

            return powerpointKey != null;
        }
    }
}
