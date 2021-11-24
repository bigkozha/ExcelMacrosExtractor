using ExcelHelper;

namespace Tester
{
    internal class Program
    {
        static void Main(string[] args)
        {
            var extractor = new ExcelMacrosExtractor();

            extractor.Save("Test.xlsm", "Test");

            var errorMessage = "";
            extractor.TrySave("Test.xlsm", "TryTest", out errorMessage);
        }
    }
}
