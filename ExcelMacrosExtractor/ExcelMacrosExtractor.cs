using OfficeOpenXml;
using OfficeOpenXml.VBA;
using System;
using System.IO;
using System.Linq;

namespace ExcelHelper
{
    public class ExcelMacrosExtractor : IExcelMacrosExtractor
    {
        private readonly string[] _allowedExtensions = new string[] { ".xlsm" };
        private readonly string classModulesFolderName = "Class Modules";
        private readonly string modulesFolderName = "Modules";
        private readonly string msExcelObjectsFolderName = "Microsoft Excel Objects";


        public void Save(string macroFilePath, string destinationFolderPath)
        {
            var fileInfo = new FileInfo(macroFilePath);

            if (!fileInfo.Exists)
            {
                throw new ExcelMacrosExtractorException($"File {macroFilePath} does not exist");
            }

            if (!_allowedExtensions.Contains(fileInfo.Extension))
            {
                throw new ExcelMacrosExtractorException($"Extension {fileInfo.Extension} is not supported");
            }

            var package = new ExcelPackage(fileInfo);

            try
            {
                foreach (var module in package.Workbook.VbaProject.Modules)
                {
                    WriteSourceCodeToFile(module, destinationFolderPath);
                }
            }
            catch (NullReferenceException)
            {
                throw new NullReferenceException($"The file ${macroFilePath} is corrupted");
            }
        }

        public bool TrySave(string macroFilePath, string destinationFolderPath, out string errorMessage)
        {
            errorMessage = string.Empty;

            try
            {
                Save(macroFilePath, destinationFolderPath);
                return true;
            }
            catch (ExcelMacrosExtractorException ex)
            {
                errorMessage = ex.Message;
                return false;
            }
            catch (NullReferenceException ex) when (ex.Message.Equals($"The file ${macroFilePath} is corrupted"))
            {
                errorMessage = $"The file ${macroFilePath} is corrupted";
                return false;
            }
        }

        private void WriteSourceCodeToFile(ExcelVBAModule module, string destinationFolderPath)
        {
            var folderName = FolderNameByType(module.Type);
            var directory = Directory.CreateDirectory(Path.Combine(destinationFolderPath, folderName));

            using (var stream = File.CreateText(Path.Combine(directory.FullName, $"{module.Name}.txt")))
            {
                stream.Write(CommentStrangeCode(module.Code));
            }
        }

        private string CommentStrangeCode(string sourceCode) => sourceCode.Replace("Attribute", "'Attribute");

        private string FolderNameByType(eModuleType type)
        {
            switch (type)
            {
                case eModuleType.Module:
                    return modulesFolderName;
                case eModuleType.Class:
                    return classModulesFolderName;
                case eModuleType.Document:
                    return msExcelObjectsFolderName;
                default:
                    return string.Empty;
            }
        }
    }
}
