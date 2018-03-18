using System;
using System.Collections;
using System.IO;
using OfficeInterop;

namespace BatchPowerPointToPDF.ConsoleApp
{
    class PowerPointConverter
    {
        private ArrayList givenFilenames;
        private PowerPointInteropLibrary exporter;

        public PowerPointConverter()
        {
            givenFilenames = new ArrayList();
            exporter = new PowerPointInteropLibrary();
            PrintWelcomeMessage();
        }

        public void ParsePaths(string[] paths)

        {
            // No arguments were given.
            if (paths.Length == 0)
            {
                PrintHelpMessage();
            }

            var givenFilenames = new ArrayList();
            for (int i = 0; i < paths.Length; i++)
            {
                try
                {
                    // Convert entire directory
                    if (paths[i] == "*")
                    {
                        var currentDirectory = Environment.CurrentDirectory;
                        AddPowerPointsInDirectory(currentDirectory);
                    }

                    // Check if file exists 
                    else if (File.Exists(paths[i]))
                    {
                        if (IsPowerPoint(Path.GetFullPath(paths[i])))
                        {
                            AddPowerPointToConvert(Path.GetFullPath(paths[i]));
                        }
                    }

                    else if (Directory.Exists(Path.GetFullPath(paths[i])))
                    {
                        AddPowerPointsInDirectory(Path.GetFullPath(paths[i]));

                    }
                }


                catch (Exception e)
                {
                    SomethingBadHappened();
                    Console.WriteLine("Here's what we know: {0}", e.Message);
                }
            }
        }

        private void AddPowerPointsInDirectory(string directory)
        {
            var allFilesInDirectory = Directory.GetFiles(directory);

            foreach (var file in allFilesInDirectory)
            {
                if (IsPowerPoint(file))
                {
                    AddPowerPointToConvert(file);
                }
            }
        }

        public void convertPowerPoints()
        {
            foreach (string powerpoint in givenFilenames)
            {
                var succeed = exporter.ConvertToPdf(powerpoint);
                if (succeed)
                {
                    Console.WriteLine("Converted: " + powerpoint);
                }
            }
        }

        private void AddPowerPointToConvert(string file)
        {
            givenFilenames.Add(Path.GetFullPath(file));
        }

        private void PrintWelcomeMessage()
        {
            Console.WriteLine("PowerPoint to PDF utlity");
            Console.WriteLine("Developed by Matthew Weihl: https://github.com/mattweihl/BatchPowerPointToPDF");
        }

        private void PrintHelpMessage()
        {
            var exeName = System.AppDomain.CurrentDomain.FriendlyName;
            Console.WriteLine("Usage: {0} paths", exeName);
            Console.WriteLine("\npaths can include: relative and full paths to files, folders, and '*' to convert all PowerPoint files in the current directory.");
            Console.WriteLine("Example: " + exeName + " example.pptx");
            Console.WriteLine("\n");
        }

        private void SomethingBadHappened()
        {
            Console.WriteLine("Something happened. Check your parameters.");
            PrintHelpMessage();
        }
        private bool IsPowerPoint(string fullPath) => (Path.GetExtension(fullPath) == ".pptx" || Path.GetExtension(fullPath) == ".ppt");
    }
}
