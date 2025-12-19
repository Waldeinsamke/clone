// FileProcessor.cs - 文件处理器
using System;
using System.IO;
using System.Linq;
using TestReportAutoFiller.Services;

namespace TestReportAutoFiller
{
    public class FileProcessor
    {
        private readonly ReportFillerService _fillerService = new ReportFillerService();

        public void ProcessSingleFile(string filePath)
        {
            // 处理可能包含空格的路径
            if (string.IsNullOrWhiteSpace(filePath))
            {
                Console.WriteLine("文件路径不能为空");
                return;
            }
            
            // 移除可能的引号和前后空格
            filePath = filePath.Trim().Trim('"').Trim('\'');
            
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"文件不存在: {filePath}");
                Console.WriteLine($"提示：如果路径包含空格，请确保路径正确，或使用引号括起来");
                return;
            }

            if (!filePath.EndsWith(".docx", StringComparison.OrdinalIgnoreCase))
            {
                Console.WriteLine("只支持.docx格式文件");
                return;
            }

            try
            {
                // 检查文件是否被占用
                if (IsFileLocked(filePath))
                {
                    Console.WriteLine("文件可能被其他程序占用，请关闭后重试");
                    return;
                }

                Console.WriteLine($"开始处理: {Path.GetFileName(filePath)}");
                _fillerService.FillWordDocument(filePath);
                Console.WriteLine($"处理完成: {Path.GetFileName(filePath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"处理失败: {ex.Message}");
            }
        }

        public void ProcessDirectory(string directoryPath)
        {
            // 处理可能包含空格的路径
            if (string.IsNullOrWhiteSpace(directoryPath))
            {
                Console.WriteLine("目录路径不能为空");
                return;
            }
            
            // 移除可能的引号和前后空格
            directoryPath = directoryPath.Trim().Trim('"').Trim('\'');
            
            if (!Directory.Exists(directoryPath))
            {
                Console.WriteLine($"目录不存在: {directoryPath}");
                Console.WriteLine($"提示：如果路径包含空格，请确保路径正确，或使用引号括起来");
                return;
            }

            var docxFiles = Directory.GetFiles(directoryPath, "*.docx");

            if (docxFiles.Length == 0)
            {
                Console.WriteLine("目录中没有找到.docx文件");
                return;
            }

            Console.WriteLine($"找到 {docxFiles.Length} 个Word文档");

            foreach (var file in docxFiles)
            {
                Console.WriteLine($"\n处理文件: {Path.GetFileName(file)}");
                ProcessSingleFile(file);
            }

            Console.WriteLine($"\n批量处理完成，共处理 {docxFiles.Length} 个文件");
        }

        private bool IsFileLocked(string filePath)
        {
            try
            {
                using (FileStream stream = File.Open(filePath, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
                {
                    stream.Close();
                }
                return false;
            }
            catch (IOException)
            {
                return true;
            }
        }
    }
}