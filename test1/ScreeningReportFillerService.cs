// Services/ScreeningReportFillerService.cs - 筛选报告填充服务
using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using TestReportAutoFiller.Models;
using TestReportAutoFiller.Configuration;

namespace TestReportAutoFiller.Services
{
    /// <summary>
    /// 筛选报告填充服务
    /// </summary>
    public class ScreeningReportFillerService
    {
        private readonly Random _random = new Random();
        
        // 存储筛选报告中每个表格的关键频点功率值（使用条件+索引作为键）
        private readonly Dictionary<TestConditionWithIndex, Dictionary<int, double>> _keyFrequencyPowers =
            new Dictionary<TestConditionWithIndex, Dictionary<int, double>>();

        /// <summary>
        /// 处理筛选报告
        /// 筛选报告有6页：第1页封面（不填充），第2-6页每页有两个表格
        /// 第2页（第1-2个表格）：填充常温试验数据
        /// 第3-5页（第3-8个表格）：填充高温试验数据
        /// 第6页（第9-10个表格）：填充常温试验数据
        /// 每个页面的第一个表格是输出功率表，第二个表格是参数汇总表
        /// </summary>
        public void ProcessScreeningReport(WordprocessingDocument doc)
        {
#pragma warning disable CS8602 // 解引用可能出现空引用。
            var allTables = doc.MainDocumentPart.Document.Body.Elements<Table>().ToList();
#pragma warning restore CS8602 // 解引用可能出现空引用。

            Console.WriteLine($"筛选报告：找到 {allTables.Count} 个表格");

            // 筛选报告应该有10个表格（5页 × 2个表格/页）
            if (allTables.Count != 10)
            {
                Console.WriteLine($"警告：筛选报告应该有10个表格（5页 × 2个表格/页），但找到 {allTables.Count} 个");
                if (allTables.Count < 10)
                {
                    Console.WriteLine("表格数量不足，将只处理找到的表格");
                }
            }

            // 按照611交付报告的逻辑处理：每页有两个表格，第一个是输出功率表，第二个是参数汇总表
            // 第2页（第1-2个表格）：填充常温试验数据
            if (allTables.Count >= 2)
            {
                Console.WriteLine("\n=== 处理第2页（第1-2个表格）- 填充常温试验数据 ===");
                ProcessScreeningTablePair(allTables[0], allTables[1], TestCondition.Normal, 1);
            }

            // 第3页（第3-4个表格）：填充高温试验数据
            if (allTables.Count >= 4)
            {
                Console.WriteLine("\n=== 处理第3页（第3-4个表格）- 填充高温试验数据 ===");
                ProcessScreeningTablePair(allTables[2], allTables[3], TestCondition.HighTemp, 2);
            }

            // 第4页（第5-6个表格）：填充高温试验数据
            if (allTables.Count >= 6)
            {
                Console.WriteLine("\n=== 处理第4页（第5-6个表格）- 填充高温试验数据 ===");
                ProcessScreeningTablePair(allTables[4], allTables[5], TestCondition.HighTemp, 3);
            }

            // 第5页（第7-8个表格）：填充高温试验数据
            if (allTables.Count >= 8)
            {
                Console.WriteLine("\n=== 处理第5页（第7-8个表格）- 填充高温试验数据 ===");
                ProcessScreeningTablePair(allTables[6], allTables[7], TestCondition.HighTemp, 4);
            }

            // 第6页（第9-10个表格）：填充常温试验数据
            if (allTables.Count >= 10)
            {
                Console.WriteLine("\n=== 处理第6页（第9-10个表格）- 填充常温试验数据 ===");
                ProcessScreeningTablePair(allTables[8], allTables[9], TestCondition.Normal, 5);
            }
        }

        /// <summary>
        /// 处理筛选报告的表格对（按照611交付报告的逻辑）
        /// 第一个表格是输出功率表，第二个表格是参数汇总表
        /// </summary>
        private void ProcessScreeningTablePair(Table outputPowerTable, Table parameterTable, TestCondition condition, int pageIndex)
        {
            // 创建唯一的条件标识（组合条件和页面索引）
            var uniqueCondition = new TestConditionWithIndex(condition, pageIndex);

            Console.WriteLine($"处理第{pageIndex}页，条件: {condition}");

            // 第一步：先处理第二个表格（参数汇总表），读取关键频点的功率值
            Console.WriteLine("第一步：读取参数汇总表的关键频点功率值...");
            ReadKeyFrequencyPowersFromParameterTable(parameterTable, uniqueCondition);

            // 第二步：使用读取到的功率值填充第一个表格（输出功率表）
            Console.WriteLine("第二步：使用关键频点功率值填充输出功率表...");
            FillOutputPowerTableUsingKeyFrequencies(outputPowerTable, uniqueCondition);

            // 第三步：填充参数汇总表的其他指标（跳过频率准确度和输出功率）
            Console.WriteLine("第三步：填充参数汇总表的其他指标...");
            FillParameterTableExceptKeyColumns(parameterTable, condition);

            Console.WriteLine($"已完成第{pageIndex}页的 {condition} 测试数据填充");
        }

        /// <summary>
        /// 处理筛选报告的单个表格
        /// 使用表格索引创建唯一的条件标识，确保每个表格的关键频点功率值独立
        /// </summary>
        private void ProcessScreeningTable(Table table, TestCondition condition, int tableIndex)
        {
            // 创建唯一的条件标识（组合条件和表格索引）
            var uniqueCondition = new TestConditionWithIndex(condition, tableIndex);
            
            var rows = table.Elements<TableRow>().ToList();
            Console.WriteLine($"表格 {tableIndex} 共有 {rows.Count} 行");
            
            // 检查表格是否包含输出功率表的结构（包含"输出功率"或"0±1.5dB"）
            bool hasOutputPowerTable = false;
            bool hasParameterTable = false;
            
            // 调试：输出前几行的内容以便诊断
            Console.WriteLine("检查表格结构，前5行的第一列内容：");
            for (int i = 0; i < Math.Min(5, rows.Count); i++)
            {
                var cells = rows[i].Elements<TableCell>().ToList();
                if (cells.Count > 0)
                {
                    string firstCellText = TableHelper.GetCellText(cells[0]).Trim();
                    Console.WriteLine($"  第{i + 1}行第一列: \"{firstCellText}\"");
                }
            }
            
            foreach (var row in rows)
            {
                var cells = row.Elements<TableCell>().ToList();
                if (cells.Count > 0)
                {
                    string firstCellText = TableHelper.GetCellText(cells[0]).Trim();
                    
                    // 检查是否是输出功率表
                    if (firstCellText.Contains("输出功率") || firstCellText.Contains("0±1.5dB"))
                    {
                        hasOutputPowerTable = true;
                        Console.WriteLine($"找到输出功率表标识: \"{firstCellText}\"");
                    }
                    
                    // 检查是否是参数汇总表（包含关键频率点1025、1101、1150）
                    if (firstCellText == "1025" || firstCellText == "1101" || firstCellText == "1150")
                    {
                        hasParameterTable = true;
                        Console.WriteLine($"找到参数汇总表标识: \"{firstCellText}\"");
                    }
                }
            }

            Console.WriteLine($"表格识别结果: hasOutputPowerTable={hasOutputPowerTable}, hasParameterTable={hasParameterTable}");

            // 如果既没有输出功率表也没有参数汇总表，尝试按611交付报告的方式处理
            // 筛选报告的表格可能结构与611交付报告类似，但可能在同一表格中
            if (!hasOutputPowerTable && !hasParameterTable)
            {
                Console.WriteLine("未识别到标准表格结构，尝试按611交付报告方式处理...");
                // 假设表格结构与611交付报告类似，尝试查找并填充
                // 先尝试查找是否有频率行（包含1025-1150范围的数字）
                bool hasFrequencyRow = false;
                foreach (var row in rows)
                {
                    var cells = row.Elements<TableCell>().ToList();
                    if (cells.Count > 0)
                    {
                        string firstCellText = TableHelper.GetCellText(cells[0]).Trim();
                        if (int.TryParse(firstCellText, out int freq) && freq >= 1025 && freq <= 1150)
                        {
                            hasFrequencyRow = true;
                            break;
                        }
                    }
                }
                
                if (hasFrequencyRow)
                {
                    Console.WriteLine("检测到频率行，假设表格包含参数汇总表结构");
                    hasParameterTable = true;
                }
                else
                {
                    // 如果连频率行都没有，假设这是一个输出功率表
                    Console.WriteLine("未检测到频率行，假设这是一个输出功率表");
                    hasOutputPowerTable = true;
                }
            }

            // 如果表格包含参数汇总表，先读取关键频点功率值
            if (hasParameterTable)
            {
                Console.WriteLine("第一步：读取参数汇总表的关键频点功率值...");
                ReadKeyFrequencyPowersFromParameterTable(table, uniqueCondition);
            }
            else
            {
                // 如果没有参数汇总表，需要为关键频点生成随机功率值
                Console.WriteLine("未找到参数汇总表，生成关键频点功率值...");
                var keyPowers = new Dictionary<int, double>();
                int[] keyFrequencies = { 1025, 1101, 1150 };
                foreach (var freq in keyFrequencies)
                {
                    double randomPower = Math.Round(
                        GenerateRandomValue(GlobalParameters.OutputPowerRange),
                        1
                    );
                    keyPowers[freq] = randomPower;
                    Console.WriteLine($"生成 {freq}MHz 的输出功率: {randomPower}");
                }
                _keyFrequencyPowers[uniqueCondition] = keyPowers;
            }

            // 如果表格包含输出功率表，填充输出功率
            if (hasOutputPowerTable)
            {
                Console.WriteLine("第二步：填充输出功率表...");
                FillOutputPowerTableUsingKeyFrequencies(table, uniqueCondition);
            }

            // 如果表格包含参数汇总表，填充其他指标
            if (hasParameterTable)
            {
                Console.WriteLine("第三步：填充参数汇总表的其他指标...");
                FillParameterTableExceptKeyColumns(table, condition);
            }

            Console.WriteLine($"已完成第{tableIndex}个表格的 {condition} 测试数据填充");
        }

        /// <summary>
        /// 从参数汇总表读取关键频点功率值（用于筛选报告）
        /// </summary>
        private void ReadKeyFrequencyPowersFromParameterTable(Table table, TestConditionWithIndex uniqueCondition)
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
                    string firstCellText = TableHelper.GetCellText(cells[0]).Trim();

                    // 检查是否为关键频率点
                    if (int.TryParse(firstCellText, out int freq) && keyFrequencies.Contains(freq))
                    {
                        // 第二列是输出功率（跳过频率准确度列）
                        if (cells.Count > 2)
                        {
                            string powerText = TableHelper.GetCellText(cells[2]).Trim();
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

            _keyFrequencyPowers[uniqueCondition] = keyPowers;
        }

        /// <summary>
        /// 使用关键频点功率值填充输出功率表（用于筛选报告）
        /// </summary>
        private void FillOutputPowerTableUsingKeyFrequencies(Table table, TestConditionWithIndex uniqueCondition)
        {
            if (!_keyFrequencyPowers.ContainsKey(uniqueCondition))
            {
                Console.WriteLine($"警告: 没有找到表格 {uniqueCondition.TableIndex} 的关键频点功率值");
                return;
            }

            var keyPowers = _keyFrequencyPowers[uniqueCondition];

            // 首先输出读取到的关键频点功率值
            Console.WriteLine($"\n=== 使用以下关键频点功率值填充表格 {uniqueCondition.TableIndex} ===");
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
            bool foundPowerRow = false;
            foreach (var row in rows)
            {
                var cells = row.Elements<TableCell>().ToList();
                if (cells.Count == 0) continue;

                string firstCellText = TableHelper.GetCellText(cells[0]).Trim();

                // 只处理输出功率行
                if (firstCellText.Contains("输出功率") || firstCellText.Contains("0±1.5dB"))
                {
                    foundPowerRow = true;
                    Console.WriteLine($"\n处理输出功率行，第一列内容: \"{firstCellText}\"，共有 {cells.Count} 列");

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
                        string cellText = TableHelper.GetCellText(cells[cellIndex]).Trim();
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

                        // 填充单元格
                        TableHelper.FillCellWithValue(cells[cellIndex], powerValue.ToString("F1"));
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
            
            if (!foundPowerRow)
            {
                Console.WriteLine("警告: 未找到输出功率行，尝试查找其他可能的输出功率行...");
                // 尝试查找包含数字的行（可能是频率行后的功率行）
                for (int rowIndex = 0; rowIndex < rows.Count; rowIndex++)
                {
                    var row = rows[rowIndex];
                    var cells = row.Elements<TableCell>().ToList();
                    if (cells.Count > 1)
                    {
                        string firstCellText = TableHelper.GetCellText(cells[0]).Trim();
                        // 如果第一列是数字（频率），下一行可能是功率行
                        if (int.TryParse(firstCellText, out int freq) && freq >= 1025 && freq <= 1150)
                        {
                            // 检查下一行
                            if (rowIndex + 1 < rows.Count)
                            {
                                var nextRow = rows[rowIndex + 1];
                                var nextCells = nextRow.Elements<TableCell>().ToList();
                                if (nextCells.Count > 1)
                                {
                                    Console.WriteLine($"找到可能的输出功率行（在频率 {freq}MHz 之后）");
                                    // 这里可以添加额外的处理逻辑
                                }
                            }
                        }
                    }
                }
            }

            Console.WriteLine($"\n=== 填充完成 ===");
            Console.WriteLine($"总填充单元格数: {totalFilled}");
            Console.WriteLine($"处理频率范围: 1025MHz - {currentFrequency - 1}MHz");
        }

        /// <summary>
        /// 填充参数汇总表的其他指标（跳过频率准确度和输出功率）
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
                    string firstCellText = TableHelper.GetCellText(cells[0]).Trim();

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
                TableHelper.FillCellWithValue(cells[3], phaseNoise1K.ToString("F1"));
            }

            // 第4列：输出相噪（10KHz）
            if (cells.Count > 4)
            {
                double phaseNoise10K = Math.Round(
                    GenerateRandomValue(GlobalParameters.PhaseNoise10KHzRange),
                    1
                );
                TableHelper.FillCellWithValue(cells[4], phaseNoise10K.ToString("F1"));
            }

            // 第5列：杂散抑制
            if (cells.Count > 5)
            {
                double spuriousSupp = Math.Round(
                    GenerateRandomValue(GlobalParameters.SpuriousSuppressionRange),
                    1
                );
                TableHelper.FillCellWithValue(cells[5], spuriousSupp.ToString("F1"));
            }

            // 第6列：谐波抑制
            if (cells.Count > 6)
            {
                double harmonicSupp = Math.Round(
                    GenerateRandomValue(GlobalParameters.HarmonicSuppressionRange),
                    1
                );
                TableHelper.FillCellWithValue(cells[6], harmonicSupp.ToString("F1"));
            }

            // 第7列：频率切换时间
            if (cells.Count > 7)
            {
                double switchTime = Math.Round(
                    GenerateRandomValue(GlobalParameters.FrequencySwitchTimeRange)
                );
                TableHelper.FillCellWithValue(cells[7], switchTime.ToString("F0"));
            }

            // 第8列：电压
            if (cells.Count > 8)
            {
                TableHelper.FillCellWithValue(cells[8], GlobalParameters.Voltage.ToString("F1"));
            }

            // 第9列：电流
            if (cells.Count > 9)
            {
                var currentRange = GetCurrentRangeForCondition(condition);
                double current = Math.Round(GenerateRandomValue(currentRange));
                TableHelper.FillCellWithValue(cells[9], current.ToString("F0"));
            }

            // 第10列：上电冲击电流
            if (cells.Count > 10)
            {
                double powerOnCurrent = Math.Round(
                    GenerateRandomValue(GlobalParameters.PowerOnCurrentRange),
                    2
                );
                TableHelper.FillCellWithValue(cells[10], powerOnCurrent.ToString("F2"));
            }

            // 第11列：冲击电流持续时间
            if (cells.Count > 11)
            {
                double impactDuration = Math.Round(
                    GenerateRandomValue(GlobalParameters.ImpactDurationRange),
                    3
                );
                TableHelper.FillCellWithValue(cells[11], impactDuration.ToString("F3"));
            }
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
        private double GenerateRandomValue((double Min, double Max) range)
        {
            return range.Min + (_random.NextDouble() * (range.Max - range.Min));
        }
    }
}
