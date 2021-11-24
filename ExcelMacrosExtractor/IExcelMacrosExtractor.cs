namespace ExcelHelper
{
    public interface IExcelMacrosExtractor
    {
        void Save(string macroFilePath, string destinationFolderPath);

        bool TrySave(string macroFilePath, string destinationFolderPath, out string errorMessage);
    }
}
