namespace BatchPowerPointToPDF.ConsoleApp
{
    class Program
    {
        static void Main(string[] args)
        {
            var converter = new PowerPointConverter();
            converter.ParsePaths(args);
            converter.convertPowerPoints();
        }
    }
}
