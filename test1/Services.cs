// Services/ReportFillerService.cs - 核心填充服务
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using TestReportAutoFiller.Models;      // 添加这行
using TestReportAutoFiller.Configuration; // 添加这行

namespace TestReportAutoFiller.Services
{
    /// <summary>
    /// 测试报告填充服务
    /// </summary>
    public class ReportFillerService
    {
        private readonly Random _random = new Random();
        private readonly Dictionary<string, object> _globalParameters = new Dictionary<string, object>();

        // 存储每个测试条件的三个关键频点功率值
        private readonly Dictionary<TestCondition, Dictionary<int, double>> _keyFrequencyPowers =
            new Dictionary<TestCondition, Dictionary<int, double>>();

        /// <summary>
        /// 填充Word文档
        /// </summary>
        public void FillWordDocument(string filePath)
        {
            try
            {
                Console.WriteLine($"开始处理文件: {Path.GetFileName(filePath)}");

                // 备份原文件
                string backupPath = CreateBackup(filePath);
                Console.WriteLine($"已备份原文件到: {backupPath}");

                using (WordprocessingDocument doc = WordprocessingDocument.Open(filePath, true))
                {
                    // 初始化全局参数（同一文档内保持一致）
                    InitializeGlobalParameters();

                    // 识别报告类型（同时传入文件名以便更准确识别）
                    var fileName = Path.GetFileName(filePath);
                    var reportType = ReportTypeIdentifier.Identify(doc, fileName);
                    Console.WriteLine($"识别到报告类型: {reportType}");

                    if (reportType == ReportType.ScreeningReport)
                    {
                        // 处理筛选报告
                        var screeningService = new ScreeningReportFillerService();
                        screeningService.ProcessScreeningReport(doc);
                    }
                    else
                    {
                        // 处理611交付报告（原有逻辑）
#pragma warning disable CS8602 // 解引用可能出现空引用。
                        var paragraphs = doc.MainDocumentPart.Document.Body.Elements<Paragraph>().ToList();
#pragma warning restore CS8602 // 解引用可能出现空引用。
                        ProcessAllTestConditions(doc, paragraphs);
                    }

                    Console.WriteLine("数据填充完成！");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"处理文件时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 初始化全局参数（同一文档内保持一致）
        /// </summary>
        private void InitializeGlobalParameters()
        {
            // 生成全局随机参数（除电流外，电流按条件变化）
            _globalParameters["OutputPower_Base"] = Math.Round(
                GenerateRandomValue(Configuration.GlobalParameters.OutputPowerRange),
                1
            );
            _globalParameters["FrequencyAccuracy"] = Math.Round(
                GenerateRandomValue(Configuration.GlobalParameters.FrequencyAccuracyRange),
                1
            );
            _globalParameters["PhaseNoise1KHz"] = Math.Round(
                GenerateRandomValue(Configuration.GlobalParameters.PhaseNoise1KHzRange),
                1
            );
            _globalParameters["PhaseNoise10KHz"] = Math.Round(
                GenerateRandomValue(Configuration.GlobalParameters.PhaseNoise10KHzRange),
                1
            );
            _globalParameters["SpuriousSuppression"] = Math.Round(
                GenerateRandomValue(Configuration.GlobalParameters.SpuriousSuppressionRange),
                1
            );
            _globalParameters["HarmonicSuppression"] = Math.Round(
                GenerateRandomValue(Configuration.GlobalParameters.HarmonicSuppressionRange),
                1
            );
            _globalParameters["FrequencySwitchTime"] = Math.Round(
                GenerateRandomValue(Configuration.GlobalParameters.FrequencySwitchTimeRange)
            );
            _globalParameters["PowerOnCurrent"] = Math.Round(
                GenerateRandomValue(Configuration.GlobalParameters.PowerOnCurrentRange),
                1
            );
            _globalParameters["ImpactDuration"] = Math.Round(
                GenerateRandomValue(Configuration.GlobalParameters.ImpactDurationRange)
            );

            // 注意：关键频率点的输出功率值在GenerateContinuousPowerValues方法中生成

            // 为关键频率点生成输出功率基准值
            foreach (var freq in Configuration.GlobalParameters.KeyFrequencies)
            {
                _globalParameters[$"OutputPower_{freq}"] = GenerateRandomValue(Configuration.GlobalParameters.OutputPowerRange);
            }
        }

        /// <summary>
        /// 处理所有测试条件（修改顺序，先处理参数表以读取数据）
        /// </summary>
        private void ProcessAllTestConditions(WordprocessingDocument doc, List<Paragraph> paragraphs)
        {
            // 扩展关键词映射，包含所有试验后条件
            var conditionKeywords = new Dictionary<string, TestCondition>
            {
                ["常温"] = TestCondition.Normal,
                ["低温"] = TestCondition.LowTemp,
                ["高温"] = TestCondition.HighTemp,
                ["功能振动"] = TestCondition.Vibration,
                ["功能冲击"] = TestCondition.Impact,
                ["低温试验后"] = TestCondition.AfterLowTemp,
                ["高温试验后"] = TestCondition.AfterHighTemp,
                ["功能振动试验后"] = TestCondition.AfterVibration,
                ["功能冲击试验后"] = TestCondition.AfterImpact
            };

            // 查找所有包含测试条件的段落
            for (int i = 0; i < paragraphs.Count; i++)
            {
                var paragraph = paragraphs[i];
                var text = GetParagraphText(paragraph);

                foreach (var keyword in conditionKeywords.Keys)
                {
                    if (text.Contains(keyword))
                    {
                        Console.WriteLine($"识别到测试条件: {keyword}");

                        // 获取测试条件
                        var condition = conditionKeywords[keyword];

                        // 查找后续的表格
                        var tables = FindTablesAfterParagraph(doc, paragraph);

                        if (tables.Count >= 2)
                        {
                            Console.WriteLine($"找到 {tables.Count} 个表格");

                            // 第一步：先处理第二个表格（参数汇总表），读取关键频点的功率值
                            if (tables.Count > 1)
                            {
                                Console.WriteLine("第一步：读取参数汇总表的关键频点功率值...");
                                ReadKeyFrequencyPowersFromParameterTable(tables[1], condition);
                            }

                            // 第二步：使用读取到的功率值填充第一个表格（输出功率表）
                            if (tables.Count > 0)
                            {
                                Console.WriteLine("第二步：使用关键频点功率值填充输出功率表...");
                                FillOutputPowerTableUsingKeyFrequencies(tables[0], condition);
                            }

                            // 第三步：填充参数汇总表的其他指标（跳过频率准确度和输出功率）
                            if (tables.Count > 1)
                            {
                                Console.WriteLine("第三步：填充参数汇总表的其他指标...");
                                FillParameterTableExceptKeyColumns(tables[1], condition);
                            }

                            Console.WriteLine($"已填充 {keyword} 测试数据");
                        }
                        else
                        {
                            Console.WriteLine($"警告: {keyword} 后只找到 {tables.Count} 个表格");
                        }
                        break;
                    }
                }
            }
        }


        /// <summary>
        /// 第一步：从参数汇总表读取三个关键频点的输出功率值（只读取，不修改）
        /// </summary>
        private void ReadKeyFrequencyPowersFromParameterTable(Table table, TestCondition condition)
        {
            var rows = table.Elements<TableRow>().ToList();
            var keyPowers = new Dictionary<int, double>();

            // 关键频率点
            int[] keyFrequencies = { 1025, 1101, 1150 };

            // 查找包含这些频率的行
            foreach (var row in rows)
            {
                var cells = row.Elements<TableCell>().ToList();
                if (cells.Count > 0)
                {
                    string firstCellText = GetCellText(cells[0]).Trim();

                    // 检查是否为关键频率点
                    if (int.TryParse(firstCellText, out int freq) && keyFrequencies.Contains(freq))
                    {
                        // 第二列是输出功率（跳过频率准确度列）
                        if (cells.Count > 2)
                        {
                            string powerText = GetCellText(cells[2]).Trim();
                            if (double.TryParse(powerText, out double power))
                            {
                                keyPowers[freq] = power;
                                Console.WriteLine($"读取到 {freq}MHz 的输出功率: {power}");
                            }
                            else
                            {
                                // 如果单元格为空或不是数字，生成一个随机值
                                double randomPower = Math.Round(
                                    GenerateRandomValue(GlobalParameters.OutputPowerRange),
                                    1
                                );
                                keyPowers[freq] = randomPower;
                                Console.WriteLine($"{freq}MHz 的输出功率为空，使用随机值: {randomPower}");
                            }
                        }
                    }
                }
            }

            // 确保三个关键频点都有值
            foreach (var freq in keyFrequencies)
            {
                if (!keyPowers.ContainsKey(freq))
                {
                    double randomPower = Math.Round(
                        GenerateRandomValue(GlobalParameters.OutputPowerRange),
                        1
                    );
                    keyPowers[freq] = randomPower;
                    Console.WriteLine($"{freq}MHz 未找到，使用随机值: {randomPower}");
                }
            }

            _keyFrequencyPowers[condition] = keyPowers;
        }

        /// <summary>
        /// 第二步：使用关键频点功率值填充输出功率表（使用新配置）
        /// </summary>
        private void FillOutputPowerTableUsingKeyFrequencies(Table table, TestCondition condition)
        {
            if (!_keyFrequencyPowers.ContainsKey(condition))
            {
                Console.WriteLine($"警告: 没有找到测试条件 {condition} 的关键频点功率值");
                return;
            }

            var keyPowers = _keyFrequencyPowers[condition];

            // 首先输出读取到的关键频点功率值
            Console.WriteLine($"\n=== 使用以下关键频点功率值填充表1 ===");
            Console.WriteLine($"1025MHz: {keyPowers[1025]:F1}dB");
            Console.WriteLine($"1101MHz: {keyPowers[1101]:F1}dB");
            Console.WriteLine($"1150MHz: {keyPowers[1150]:F1}dB");

            // 获取表格中的所有行
            var rows = table.Elements<TableRow>().ToList();

            if (rows.Count < 2) return;

            int currentFrequency = 1025;
            int totalFilled = 0;

            Console.WriteLine($"\n开始填充，起始频率: {currentFrequency}MHz");

            // 遍历每一行，只处理输出功率行
            foreach (var row in rows)
            {
                var cells = row.Elements<TableCell>().ToList();
                if (cells.Count == 0) continue;

                string firstCellText = GetCellText(cells[0]).Trim();

                // 只处理输出功率行
                if (firstCellText.Contains("输出功率") || firstCellText.Contains("0±1.5dB"))
                {
                    Console.WriteLine($"\n处理输出功率行");

                    // 填充这一行的所有单元格（跳过第一个标题单元格）
                    for (int cellIndex = 1; cellIndex < cells.Count; cellIndex++)
                    {
                        // 检查是否已经处理完所有频率点
                        if (currentFrequency > 1150)
                        {
                            Console.WriteLine("已达到最大频率 1150MHz，停止填充");
                            break;
                        }

                        // 跳过特殊符号
                        string cellText = GetCellText(cells[cellIndex]).Trim();
                        if (cellText == "/" || cellText == "／" || cellText.Contains("/"))
                        {
                            currentFrequency++;
                            continue;
                        }

                        // 使用新配置获取功率设置
                        var settings = GlobalParameters.GetPowerSettingsForFrequency(currentFrequency);
                        int keyFreq = settings.KeyFrequency;
                        double offset = settings.Offset;

                        // 计算功率值
                        double powerValue = keyPowers[keyFreq] + offset;
                        powerValue = Math.Round(powerValue, 1);

                        // 确保在合理范围内
                        powerValue = Math.Max(powerValue, -1.5);
                        powerValue = Math.Min(powerValue, 1.5);

                        FillCellWithValue(cells[cellIndex], powerValue.ToString("F1"));
                        totalFilled++;

                        // 输出关键点的填充信息
                        if (offset != 0 || currentFrequency == 1025 || currentFrequency == 1101 ||
                            currentFrequency == 1150 || currentFrequency == 1054 ||
                            currentFrequency == 1055 || currentFrequency == 1100 ||
                            currentFrequency == 1127 || currentFrequency == 1128)
                        {
                            Console.WriteLine($"  频率 {currentFrequency}MHz → 使用 {keyFreq}MHz的功率值 + {offset:F1}dB = {powerValue:F1}dB");
                        }

                        currentFrequency++;
                    }
                }
            }

            Console.WriteLine($"\n=== 填充完成 ===");
            Console.WriteLine($"总填充单元格数: {totalFilled}");
            Console.WriteLine($"处理频率范围: 1025MHz - {currentFrequency - 1}MHz");
        }

        /// <summary>
        /// 验证新的分段填充规则
        /// </summary>
        private void ValidateFillResultsWithNewRules(Dictionary<int, double> keyPowers)
        {
            Console.WriteLine("\n=== 验证新的分段填充规则 ===");
            Console.WriteLine("新的分段规则:");
            Console.WriteLine("1025-1054MHz → 使用1025MHz的功率值");
            Console.WriteLine("1055-1100MHz → 使用1025MHz的功率值 + 0.1dB");
            Console.WriteLine("1101MHz     → 使用1101MHz的功率值");
            Console.WriteLine("1102-1127MHz → 使用1101MHz的功率值");
            Console.WriteLine("1128-1150MHz → 使用1150MHz的功率值");

            Console.WriteLine("\n关键频率点功率值:");
            Console.WriteLine($"1025MHz: {keyPowers[1025]:F1}dB");
            Console.WriteLine($"1101MHz: {keyPowers[1101]:F1}dB");
            Console.WriteLine($"1150MHz: {keyPowers[1150]:F1}dB");
            Console.WriteLine($"1055-1100MHz使用: {keyPowers[1025] + 0.1:F1}dB");

            // 测试几个关键边界点
            int[] testPoints = { 1025, 1054, 1055, 1100, 1101, 1102, 1127, 1128, 1150 };

            Console.WriteLine("\n边界点验证:");
            foreach (int freq in testPoints)
            {
                double powerValue;
                if (freq <= 1054)
                    powerValue = keyPowers[1025];
                else if (freq >= 1055 && freq <= 1100)
                    powerValue = keyPowers[1025] + 0.1;
                else if (freq == 1101)
                    powerValue = keyPowers[1101];
                else if (freq >= 1102 && freq <= 1127)
                    powerValue = keyPowers[1101];
                else // freq >= 1128
                    powerValue = keyPowers[1150];

                powerValue = Math.Round(powerValue, 1);

                Console.WriteLine($"{freq}MHz → 使用功率值: {powerValue:F1}dB");
            }
        }

        /// <summary>
        /// 根据频率获取对应的关键频率点（独立方法，不依赖可能错误的全局配置）
        /// </summary>
        private int GetKeyFrequencyForFrequency(int frequency)
        {
            // 按照您的分段规则：
            // 1025-1065: 使用 1025
            // 1066-1101: 使用 1101
            // 1102-1127: 使用 1101
            // 1128-1150: 使用 1150

            if (frequency >= 1025 && frequency <= 1065)
                return 1025;
            else if (frequency >= 1066 && frequency <= 1127)
                return 1101;
            else if (frequency >= 1128 && frequency <= 1150)
                return 1150;
            else
            {
                Console.WriteLine($"警告: 频率 {frequency}MHz 不在任何分段中，使用1025MHz");
                return 1025;
            }
        }


        /// <summary>
        /// 获取表格中的频率点数量（根据实际文档结构）
        /// </summary>
        private int GetFrequenciesPerRow(Table table)
        {
            // 根据示例文档，每行有23个频率点
            // 您可以根据实际文档结构调整这个值
            return 23;
        }


        /// <summary>
        /// 第三步：填充参数汇总表的其他指标（第二个表格，跳过频率准确度和输出功率）
        /// </summary>
        private void FillParameterTableExceptKeyColumns(Table table, TestCondition condition)
        {
            var rows = table.Elements<TableRow>().ToList();

            // 查找数据行（包含关键频率点的行）
            List<TableRow> dataRows = new List<TableRow>();

            for (int i = 0; i < rows.Count; i++)
            {
                var row = rows[i];
                var cells = row.Elements<TableCell>().ToList();

                if (cells.Count > 0)
                {
                    string firstCellText = GetCellText(cells[0]).Trim();

                    // 检查是否是频率行（包含1025、1101、1150）
                    if (firstCellText == "1025" || firstCellText == "1101" || firstCellText == "1150")
                    {
                        dataRows.Add(row);
                    }
                }
            }

            // 填充三个关键频率点的数据（跳过频率准确度和输出功率）
            for (int i = 0; i < Math.Min(3, dataRows.Count); i++)
            {
                var row = dataRows[i];
                var cells = row.Elements<TableCell>().ToList();

                if (cells.Count >= 11)
                {
                    // 只填充第4列及以后的指标（跳过频率准确度和输出功率）
                    FillParameterRowExceptKeyColumns(cells, condition);
                }
            }
        }

        /// <summary>
        /// 填充参数行（跳过频率准确度和输出功率）
        /// </summary>
        private void FillParameterRowExceptKeyColumns(List<TableCell> cells, TestCondition condition)
        {
            if (cells.Count < 11) return;

            // 第0列：频率（已存在）
            // 第1列：频率准确度（跳过，保留用户数据）
            // 第2列：输出功率（跳过，保留用户数据）

            // 第3列：输出相噪（1KHz）
            if (cells.Count > 3)
            {
                double phaseNoise1K = Math.Round(
                    GenerateRandomValue(GlobalParameters.PhaseNoise1KHzRange),
                    1
                );
                FillCellWithValue(cells[3], phaseNoise1K.ToString("F1"));
            }

            // 第4列：输出相噪（10KHz）
            if (cells.Count > 4)
            {
                double phaseNoise10K = Math.Round(
                    GenerateRandomValue(GlobalParameters.PhaseNoise10KHzRange),
                    1
                );
                FillCellWithValue(cells[4], phaseNoise10K.ToString("F1"));
            }

            // 第5列：杂散抑制
            if (cells.Count > 5)
            {
                double spuriousSupp = Math.Round(
                    GenerateRandomValue(GlobalParameters.SpuriousSuppressionRange),
                    1
                );
                FillCellWithValue(cells[5], spuriousSupp.ToString("F1"));
            }

            // 第6列：谐波抑制
            if (cells.Count > 6)
            {
                double harmonicSupp = Math.Round(
                    GenerateRandomValue(GlobalParameters.HarmonicSuppressionRange),
                    1
                );
                FillCellWithValue(cells[6], harmonicSupp.ToString("F1"));
            }

            // 第7列：频率切换时间
            if (cells.Count > 7)
            {
                double switchTime = Math.Round(
                    GenerateRandomValue(GlobalParameters.FrequencySwitchTimeRange)
                );
                FillCellWithValue(cells[7], switchTime.ToString("F0"));
            }

            // 第8列：电压
            if (cells.Count > 8)
            {
                FillCellWithValue(cells[8], GlobalParameters.Voltage.ToString("F1"));
            }

            // 第9列：电流
            if (cells.Count > 9)
            {
                var currentRange = GetCurrentRangeForCondition(condition);
                double current = Math.Round(GenerateRandomValue(currentRange));
                FillCellWithValue(cells[9], current.ToString("F0"));
            }

            // 第10列：上电冲击电流
            if (cells.Count > 10)
            {
                double powerOnCurrent = Math.Round(
                    GenerateRandomValue(GlobalParameters.PowerOnCurrentRange),2
                                    );
                FillCellWithValue(cells[10], powerOnCurrent.ToString("F2"));
            }

            // 第11列：冲击电流持续时间
            if (cells.Count > 11)
            {
                double impactDuration = Math.Round(
                    GenerateRandomValue(GlobalParameters.ImpactDurationRange),3
                );
                FillCellWithValue(cells[11], impactDuration.ToString("F3"));
            }
        }

        /// <summary>
        /// 填充输出功率表
        /// </summary>
        private void FillOutputPowerTable(Table table, Models.TestCondition condition)
        {
            // 生成输出功率序列（保持连续性）
            var powerValues = GenerateContinuousPowerValues();

            // 获取表格中的所有行
            var rows = table.Elements<TableRow>().ToList();

            if (rows.Count < 2) return;

            // 第一行是表头（频率MHz值），从第二行开始
            // 表格结构通常是：频率行 -> 输出功率行 -> 频率行 -> 输出功率行 -> ...

            int freqIndex = 0; // 当前处理的频率点索引
            bool isPowerRow = false; // 当前行是否为输出功率行

            // 遍历每一行（跳过可能的表头）
            for (int rowIndex = 0; rowIndex < rows.Count; rowIndex++)
            {
                var row = rows[rowIndex];
                var cells = row.Elements<TableCell>().ToList();

                if (cells.Count == 0) continue;

                // 检查这一行是否是输出功率行
                // 输出功率行的特征：第一个单元格包含"输出功率"或"0±1.5dB"
                string firstCellText = GetCellText(cells[0]).Trim();
                isPowerRow = firstCellText.Contains("输出功率") || firstCellText.Contains("0±1.5dB");

                if (isPowerRow)
                {
                    // 这是输出功率行，填充数据
                    // 跳过第一列（包含"输出功率"或"0±1.5dB"）
                    for (int cellIndex = 1; cellIndex < cells.Count; cellIndex++)
                    {
                        // 检查这个单元格是否有"/"符号，如果有则不填充
                        string cellText = GetCellText(cells[cellIndex]).Trim();
                        if (cellText == "/" || cellText == "／" || cellText.Contains("/"))
                        {
                            continue; // 跳过不填充的单元格
                        }

                        if (freqIndex < powerValues.Length)
                        {
                            // 填充数据，保留一位小数
                            FillCellWithValue(cells[cellIndex], powerValues[freqIndex].ToString("F1"));
                            freqIndex++;
                        }
                        else
                        {
                            break; // 所有频率点已填充
                        }
                    }
                }
                else
                {
                    // 这是频率行，不要填充任何数据
                    // 频率行可能已经包含频率值（如1025MHz），我们保持原样
                    continue;
                }

                if (freqIndex >= powerValues.Length)
                {
                    break; // 所有频率点已填充
                }
            }
        }

        /// <summary>
        /// 填充参数汇总表（修正版，处理合并单元格）
        /// </summary>
        private void FillParameterTable(Table table, TestCondition condition)
        {
            var rows = table.Elements<TableRow>().ToList();

            if (rows.Count < 4) return; // 至少需要表头+3行数据

            // 第一行是表头，包含条件说明
            // 第二行可能是条件行（输出相噪的条件）或空白行
            // 我们需要找到真正的数据行（包含1025、1101、1150的行）

            List<TableRow> dataRows = new List<TableRow>();
            List<int> frequencies = new List<int>();

            // 查找数据行（包含关键频率点的行）
            for (int i = 0; i < rows.Count; i++)
            {
                var row = rows[i];
                var cells = row.Elements<TableCell>().ToList();

                if (cells.Count > 0)
                {
                    string firstCellText = GetCellText(cells[0]).Trim();

                    // 检查是否是频率行（包含1025、1101、1150）
                    if (firstCellText == "1025" || firstCellText == "1101" || firstCellText == "1150")
                    {
                        dataRows.Add(row);

                        // 记录频率值
                        if (int.TryParse(firstCellText, out int freq))
                        {
                            frequencies.Add(freq);
                        }
                    }
                }
            }

            // 确保找到了3个数据行
            if (dataRows.Count >= 3)
            {
                // 填充三个关键频率点的数据
                for (int i = 0; i < 3; i++)
                {
                    var row = dataRows[i];
                    var cells = row.Elements<TableCell>().ToList();

                    if (cells.Count >= 11)
                    {
                        FillParameterRow(cells, frequencies[i], condition);
                    }
                    else
                    {
                        Console.WriteLine($"警告: 第{i + 1}个数据行单元格数量不足: {cells.Count}");
                    }
                }
            }
            else
            {
                Console.WriteLine($"警告: 找到的数据行数量不足: {dataRows.Count}，将尝试填充前3行");

                // 备用方案：填充表格的前3行（跳过可能的表头行）
                int startRow = FindDataRowStartIndex(rows);
                for (int i = 0; i < 3 && (startRow + i) < rows.Count; i++)
                {
                    var row = rows[startRow + i];
                    var cells = row.Elements<TableCell>().ToList();

                    if (cells.Count >= 11)
                    {
                        FillParameterRow(cells, Configuration.GlobalParameters.KeyFrequencies[i], condition);
                    }
                }
            }
        }

        /// <summary>
        /// 查找数据行开始索引
        /// </summary>
        private int FindDataRowStartIndex(List<TableRow> rows)
        {
            // 跳过表头和条件行，找到第一个可能是数据行的行
            for (int i = 0; i < Math.Min(5, rows.Count); i++) // 检查前5行
            {
                var row = rows[i];
                var cells = row.Elements<TableCell>().ToList();

                if (cells.Count > 0)
                {
                    string firstCellText = GetCellText(cells[0]).Trim();

                    // 如果第一列是数字，很可能是数据行
                    if (int.TryParse(firstCellText, out int num))
                    {
                        return i;
                    }

                    // 如果第一列是"频率"，这是表头，跳过
                    if (firstCellText.Contains("频率"))
                    {
                        continue;
                    }
                }
            }

            return 2; // 默认从第三行开始（跳过表头和条件行）
        }


        /// <summary>
        /// 填充参数行（修正版，处理合并单元格）
        /// </summary>
        private void FillParameterRow(List<TableCell> cells, int frequency, TestCondition condition)
        {
            if (cells.Count < 11)
            {
                Console.WriteLine($"警告: 单元格数量不足: {cells.Count}");
                return;
            }

            // 注意：由于有合并单元格，列索引可能与预期不同
            // 我们需要根据内容判断列的位置

            // 第0列：频率（应该已经填好了）

            // 第1列：频率准确度
            if (cells.Count > 1)
            {
                double freqAccuracy = (double)_globalParameters["FrequencyAccuracy"];
                FillCellWithValue(cells[1], freqAccuracy.ToString("F1"));
            }

            // 第2列：输出功率（使用对应频率的功率值）
            if (cells.Count > 2)
            {
                string powerKey = $"OutputPower_{frequency}";
                double outputPower = _globalParameters.ContainsKey(powerKey)
                    ? (double)_globalParameters[powerKey]
                    : (double)_globalParameters["OutputPower_Base"];
                FillCellWithValue(cells[2], outputPower.ToString("F1"));
            }

            // 第3列：输出相噪（1KHz）
            if (cells.Count > 3)
            {
                double phaseNoise1K = (double)_globalParameters["PhaseNoise1KHz"];
                FillCellWithValue(cells[3], phaseNoise1K.ToString("F1"));
            }

            // 第4列：输出相噪（10KHz）
            if (cells.Count > 4)
            {
                double phaseNoise10K = (double)_globalParameters["PhaseNoise10KHz"];
                FillCellWithValue(cells[4], phaseNoise10K.ToString("F1"));
            }

            // 第5列：杂散抑制
            if (cells.Count > 5)
            {
                double spuriousSupp = (double)_globalParameters["SpuriousSuppression"];
                FillCellWithValue(cells[5], spuriousSupp.ToString("F1"));
            }

            // 第6列：谐波抑制
            if (cells.Count > 6)
            {
                double harmonicSupp = (double)_globalParameters["HarmonicSuppression"];
                FillCellWithValue(cells[6], harmonicSupp.ToString("F1"));
            }

            // 第7列：频率切换时间
            if (cells.Count > 7)
            {
                double switchTime = Math.Round((double)_globalParameters["FrequencySwitchTime"]);
                FillCellWithValue(cells[7], switchTime.ToString("F0"));
            }

            // 第8列：电压
            if (cells.Count > 8)
            {
                FillCellWithValue(cells[8], Configuration.GlobalParameters.Voltage.ToString("F1"));
            }

            // 第9列：电流
            if (cells.Count > 9)
            {
                var currentRange = GetCurrentRangeForCondition(condition);
                double current = Math.Round(GenerateRandomValue(currentRange));
                FillCellWithValue(cells[9], current.ToString("F0"));
            }

            // 第10列：上电冲击电流
            if (cells.Count > 10)
            {
                double powerOnCurrent = (double)_globalParameters["PowerOnCurrent"];
                FillCellWithValue(cells[10], powerOnCurrent.ToString("F1"));
            }

            // 第11列：冲击电流持续时间
            if (cells.Count > 11)
            {
                double impactDuration = Math.Round((double)_globalParameters["ImpactDuration"]);
                FillCellWithValue(cells[11], impactDuration.ToString("F0"));
            }
        }

        /// <summary>
        /// 生成连续的输出功率值（修正版）
        /// </summary>
        private double[] GenerateContinuousPowerValues()
        {
            int count = GlobalParameters.FrequencyRange.End - GlobalParameters.FrequencyRange.Start + 1;
            var values = new double[count];

            // 为关键频率点设置基准值
            var keyFreqIndices = GlobalParameters.KeyFrequencies
                .Select(f => f - GlobalParameters.FrequencyRange.Start)
                .ToArray();

            // 为关键频率点生成随机值，保留一位小数
            for (int i = 0; i < keyFreqIndices.Length; i++)
            {
                double baseValue = Math.Round(
                    GenerateRandomValue(GlobalParameters.OutputPowerRange),
                    1
                );
                _globalParameters[$"OutputPower_{GlobalParameters.KeyFrequencies[i]}"] = baseValue;
                values[keyFreqIndices[i]] = baseValue;
            }

            // 线性插值填充中间值
            for (int i = 0; i < keyFreqIndices.Length - 1; i++)
            {
                int startIdx = keyFreqIndices[i];
                int endIdx = keyFreqIndices[i + 1];
                double startValue = values[startIdx];
                double endValue = values[endIdx];

                // 计算每个中间点的值
                for (int j = startIdx + 1; j < endIdx; j++)
                {
                    double ratio = (double)(j - startIdx) / (endIdx - startIdx);
                    double interpolatedValue = startValue + (endValue - startValue) * ratio;

                    // 添加微小随机波动（不超过0.1dB），然后四舍五入保留一位小数
                    double variation = (_random.NextDouble() - 0.5) * GlobalParameters.MaxPowerChangeStep;
                    double finalValue = Math.Round(interpolatedValue + variation, 1);

                    // 确保在范围内
                    finalValue = Math.Max(finalValue, GlobalParameters.OutputPowerRange.Min);
                    finalValue = Math.Min(finalValue, GlobalParameters.OutputPowerRange.Max);

                    values[j] = finalValue;
                }
            }

            return values;
        }

        /// <summary>
        /// 获取单元格文本
        /// </summary>
        private string GetCellText(TableCell cell)
        {
            return TableHelper.GetCellText(cell);
        }

        /// <summary>
        /// 根据测试条件获取电流范围
        /// 规则：除了高温和低温外，其他所有条件都使用常温电流范围
        /// </summary>
        private (double Min, double Max) GetCurrentRangeForCondition(TestCondition condition)
        {
            // 将枚举值映射到配置字典中的键
            string conditionKey = condition switch
            {
                TestCondition.Normal => "常温",
                TestCondition.LowTemp => "低温",
                TestCondition.HighTemp => "高温",
                TestCondition.Vibration => "振动",
                TestCondition.Impact => "冲击",
                TestCondition.AfterLowTemp => "低温试验后",
                TestCondition.AfterHighTemp => "高温试验后",
                TestCondition.AfterVibration => "功能振动试验后",
                TestCondition.AfterImpact => "功能冲击试验后",
                _ => "常温" // 默认使用常温
            };

            if (GlobalParameters.CurrentConfig.ContainsKey(conditionKey))
            {
                return GlobalParameters.CurrentConfig[conditionKey];
            }

            // 如果配置中没有找到，使用常温范围作为默认值
            Console.WriteLine($"警告: 未找到条件 '{conditionKey}' 的电流配置，使用常温范围");
            return GlobalParameters.CurrentConfig["常温"];
        }

        /// <summary>
        /// 生成指定范围内的随机值
        /// </summary>
        //private double GenerateRandomValue((double Min, double Max) range)
        //{
        //    double value = range.Min + (_random.NextDouble() * (range.Max - range.Min));
        //    return Math.Round(value, 1); // 保留一位小数
        //}
        private double GenerateRandomValue((double Min, double Max) range)
        {
            return range.Min + (_random.NextDouble() * (range.Max - range.Min));
        }

        /// <summary>
        /// 创建文件备份
        /// </summary>
        private string CreateBackup(string originalPath)
        {
#pragma warning disable CS8604 // 引用类型参数可能为 null。
            string backupDir = Path.Combine(Path.GetDirectoryName(originalPath), "Backup");
#pragma warning restore CS8604 // 引用类型参数可能为 null。
            Directory.CreateDirectory(backupDir);

            string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            string backupName = $"{Path.GetFileNameWithoutExtension(originalPath)}_{timestamp}{Path.GetExtension(originalPath)}";
            string backupPath = Path.Combine(backupDir, backupName);

            File.Copy(originalPath, backupPath, true);
            return backupPath;
        }

        // 辅助方法
        private string GetParagraphText(Paragraph paragraph)
        {
            return string.Join("", paragraph.Descendants<Text>().Select(t => t.Text));
        }

        /// <summary>
        /// 填充单元格值
        /// </summary>
        private void FillCellWithValue(TableCell cell, string value)
        {
            TableHelper.FillCellWithValue(cell, value);
        }


        /// <summary>
        /// 查找段落后的表格（改进版，避免找到太远的表格）
        /// </summary>
        private List<Table> FindTablesAfterParagraph(WordprocessingDocument doc, Paragraph startParagraph)
        {
            var tables = new List<Table>();
#pragma warning disable CS8602 // 解引用可能出现空引用。
            var bodyElements = doc.MainDocumentPart.Document.Body.Elements().ToList();
#pragma warning restore CS8602 // 解引用可能出现空引用。

            int startIndex = bodyElements.IndexOf(startParagraph);
            if (startIndex == -1) return tables;

            // 只查找紧接着的表格（最多查找10个元素距离）
            int maxSearch = Math.Min(startIndex + 10, bodyElements.Count);
            for (int i = startIndex + 1; i < maxSearch; i++)
            {
                if (bodyElements[i] is Table table)
                {
                    tables.Add(table);

                    // 收集最多3个表格
                    if (tables.Count >= 3) break;
                }
                else if (bodyElements[i] is Paragraph paragraph)
                {
                    var text = GetParagraphText(paragraph);

                    // 如果遇到新的测试条件段落，停止收集
                    if (text.Contains("测试记录") || text.Contains("试验后测试") ||
                        text.Contains("常温") || text.Contains("低温") ||
                        text.Contains("高温") || text.Contains("振动") ||
                        text.Contains("冲击"))
                    {
                        // 检查是否紧接的段落还是标题的一部分
                        if (i > startIndex + 3) // 如果距离较远，可能是新的测试
                        {
                            break;
                        }
                    }
                }
            }

            return tables;
        }


    }
}