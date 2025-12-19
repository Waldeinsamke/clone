// Services/TableHelper.cs - 表格处理辅助工具类
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using TestReportAutoFiller.Configuration;

namespace TestReportAutoFiller.Services
{
    /// <summary>
    /// 表格处理辅助工具类
    /// </summary>
    public static class TableHelper
    {
        /// <summary>
        /// 获取单元格文本
        /// </summary>
        public static string GetCellText(TableCell cell)
        {
            var paragraphs = cell.Elements<Paragraph>().ToList();
            if (paragraphs.Count > 0)
            {
                var runs = paragraphs[0].Elements<Run>().ToList();
                if (runs.Count > 0)
                {
                    var texts = runs[0].Elements<Text>().ToList();
                    if (texts.Count > 0)
                    {
                        return texts[0].Text;
                    }
                }
            }
            return "";
        }

        /// <summary>
        /// 填充单元格值
        /// </summary>
        public static void FillCellWithValue(TableCell cell, string value)
        {
            var paragraphs = cell.Elements<Paragraph>().ToList();
            if (paragraphs.Count > 0)
            {
                var runs = paragraphs[0].Elements<Run>().ToList();
                if (runs.Count > 0)
                {
                    var texts = runs[0].Elements<Text>().ToList();
                    if (texts.Count > 0)
                    {
                        // 更新现有文本
                        texts[0].Text = value;
                    }
                    else
                    {
                        // 添加新文本
                        runs[0].AppendChild(new Text(value));
                    }
                }
                else
                {
                    // 添加新的Run和Text
                    var run = new Run();
                    run.AppendChild(new Text(value));
                    paragraphs[0].AppendChild(run);
                }
            }
            else
            {
                // 添加新的Paragraph、Run和Text
                var paragraph = new Paragraph();
                var run = new Run();
                run.AppendChild(new Text(value));
                paragraph.AppendChild(run);
                cell.AppendChild(paragraph);
            }
        }
    }
}
