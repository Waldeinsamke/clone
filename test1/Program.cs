using System;
using System.IO;
using System.Text;

namespace TestReportAutoFiller
{
    class Program
    {
        static void Main(string[] args)
        {

            Console.OutputEncoding = Encoding.UTF8;
            Console.InputEncoding  = Encoding.UTF8;

            ShowBanner();

            var processor = new FileProcessor();

            try
            {
                if (args.Length > 0)
                {
                    // 命令行参数方式（支持拖拽）
                    // 处理可能包含空格的路径（如果参数被空格分割，需要重新组合）
                    string fullPath = string.Join(" ", args);
                    
                    // 首先尝试作为完整路径
                    if (File.Exists(fullPath))
                    {
                        processor.ProcessSingleFile(fullPath);
                    }
                    else if (Directory.Exists(fullPath))
                    {
                        processor.ProcessDirectory(fullPath);
                    }
                    else
                    {
                        // 如果完整路径不存在，尝试逐个处理参数
                        foreach (var arg in args)
                        {
                            if (File.Exists(arg))
                            {
                                processor.ProcessSingleFile(arg);
                            }
                            else if (Directory.Exists(arg))
                            {
                                processor.ProcessDirectory(arg);
                            }
                            else
                            {
                                Console.WriteLine($"路径不存在: {arg}");
                            }
                        }
                    }
                }
                else
                {
                    // 交互式方式
                    RunInteractiveMode(processor);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"程序运行错误: {ex.Message}");
                Console.WriteLine($"详细信息: {ex.StackTrace}");
            }

            Console.WriteLine("\n按任意键退出...");
            Console.ReadKey();
        }

        static void ShowBanner()
        {
            Console.WriteLine("==========================================");
            Console.WriteLine("    A611项目测试报告自动填充系统 v1.0");
            Console.WriteLine("==========================================");
            Console.WriteLine("功能说明:");
            Console.WriteLine("  • 自动填充频率合成器测试数据");
            Console.WriteLine("  • 支持常温、低温、高温、振动、冲击条件");
            Console.WriteLine("  • 数据符合技术规范要求");
            Console.WriteLine("  • 自动备份原文件");
            Console.WriteLine("==========================================\n");
        }

        static void RunInteractiveMode(FileProcessor processor)
        {
            Console.WriteLine("请选择操作:");
            Console.WriteLine("1. 处理单个文件");
            Console.WriteLine("2. 处理当前目录所有文件");
            Console.WriteLine("3. 处理指定目录");
            Console.Write("请选择 (1-3): ");

            var choice = Console.ReadLine();

            switch (choice)
            {
                case "1":
                    Console.Write("请输入Word文档完整路径（如果路径包含空格，请用引号括起来）: ");
                    var filePath = Console.ReadLine();
                    if (string.IsNullOrWhiteSpace(filePath))
                    {
                        Console.WriteLine("路径不能为空");
                        break;
                    }
                    // 移除可能的引号（用户可能输入了引号）
                    filePath = filePath.Trim().Trim('"').Trim('\'');
#pragma warning disable CS8604 // 引用类型参数可能为 null。
                    processor.ProcessSingleFile(filePath);
#pragma warning restore CS8604 // 引用类型参数可能为 null。
                    break;

                case "2":
                    processor.ProcessDirectory(Directory.GetCurrentDirectory());
                    break;

                case "3":
                    Console.Write("请输入目录路径（如果路径包含空格，请用引号括起来）: ");
                    var dirPath = Console.ReadLine();
                    if (string.IsNullOrWhiteSpace(dirPath))
                    {
                        Console.WriteLine("路径不能为空");
                        break;
                    }
                    // 移除可能的引号（用户可能输入了引号）
                    dirPath = dirPath.Trim().Trim('"').Trim('\'');
#pragma warning disable CS8604 // 引用类型参数可能为 null。
                    processor.ProcessDirectory(dirPath);
#pragma warning restore CS8604 // 引用类型参数可能为 null。
                    break;

                default:
                    Console.WriteLine("无效选择");
                    break;
            }
        }
    }
}