// Services/ReportTypeIdentifier.cs - 报告类型识别器
using System;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using TestReportAutoFiller.Models;

namespace TestReportAutoFiller.Services
{
    /// <summary>
    /// 报告类型识别器
    /// </summary>
    public static class ReportTypeIdentifier
    {
        /// <summary>
        /// 识别报告类型
        /// </summary>
        public static ReportType Identify(WordprocessingDocument doc, string fileName = "")
        {
#pragma warning disable CS8602 // 解引用可能出现空引用。
            var paragraphs = doc.MainDocumentPart.Document.Body.Elements<Paragraph>().ToList();
#pragma warning restore CS8602 // 解引用可能出现空引用。

            // 首先检查文件名是否包含"筛选"关键词（忽略大小写和空格）
            if (!string.IsNullOrEmpty(fileName))
            {
                // 直接检查原始文件名（Contains方法可以处理包含空格的文件名）
                // 例如："A611 环境筛选 202502021.docx" 中的 "环境筛选" 可以被正确识别
                if (fileName.Contains("筛选") || fileName.Contains("环境筛选"))
                {
                    Console.WriteLine($"根据文件名识别为筛选报告: {fileName}");
                    return ReportType.ScreeningReport;
                }
            }

            // 检查文档中是否包含"筛选报告"关键词（包括"环境筛选报告"、"筛选报告"等）
            foreach (var paragraph in paragraphs)
            {
                var text = GetParagraphText(paragraph);
                if (text.Contains("筛选报告") || text.Contains("环境筛选") || text.Contains("筛选"))
                {
                    Console.WriteLine($"检测到筛选报告关键词: \"{text}\"");
                    return ReportType.ScreeningReport;
                }
            }

            // 检查表格数量：筛选报告应该有10个表格（5页 × 2个表格/页，跳过封面）
            var allTables = doc.MainDocumentPart.Document.Body.Elements<Table>().ToList();
            if (allTables.Count == 10)
            {
                Console.WriteLine($"检测到10个表格，可能是筛选报告");
                return ReportType.ScreeningReport;
            }
            // 也支持5个表格的情况（可能是旧版本格式）
            else if (allTables.Count == 5)
            {
                Console.WriteLine($"检测到5个表格，可能是筛选报告（旧版本格式）");
                return ReportType.ScreeningReport;
            }

            return ReportType.DeliveryReport611;
        }

        /// <summary>
        /// 获取段落文本
        /// </summary>
        private static string GetParagraphText(Paragraph paragraph)
        {
            return string.Join("", paragraph.Descendants<Text>().Select(t => t.Text));
        }
    }
}
