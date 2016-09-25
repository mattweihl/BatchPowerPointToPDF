using Microsoft.Office.Core;
using Microsoft.Win32;
using System;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PPTXExporterLibrary
{
    public class PPTXExporter
    {
        public static void ConvertToPDF(string pptFilename)
        {
            try
            {
                if (!OfficeInstalled())
                {
                    throw new NotSupportedException("Office is not installed.");
                }

                // Sanitizing filename strings
                pptFilename = pptFilename.Replace(@"\\", @"\");

                PowerPoint.Application app = new PowerPoint.Application();
                var presentation = app.Presentations;

                // Opening PowerPoint
                var file = app.Presentations.Open(pptFilename, MsoTriState.msoTrue, MsoTriState.msoTrue, MsoTriState.msoFalse);

                // Converting to PDF
                file.ExportAsFixedFormat(pptFilename + ".pdf", PowerPoint.PpFixedFormatType.ppFixedFormatTypePDF);
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

        public static bool OfficeInstalled()
        {
            RegistryKey powerpointKey = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\powerpnt.exe");

            if (powerpointKey != null)
            {
                powerpointKey.Close();
            }

            return powerpointKey != null;
        }
    }
}
