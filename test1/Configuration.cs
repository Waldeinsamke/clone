using System;
using System.Collections.Generic;

namespace TestReportAutoFiller.Configuration
{
    /// <summary>
    /// 全局测试参数配置
    /// </summary>
    public static class GlobalParameters
    {
        // 用户可修改区域 - 用括号标注
        // ============================================
        // [用户可修改开始]

        /// <summary>
        /// 电流值配置 - 按测试条件变化
        /// 要求：除了高温和低温外，其他所有试验（包括试验后）都使用常温电流范围
        /// </summary>
        public static Dictionary<string, (double Min, double Max)> CurrentConfig = new Dictionary<string, (double, double)>
        {
            ["常温"] = (195, 200),      // 常温电流范围
            ["高温"] = (220, 225),      // 高温电流范围
            ["低温"] = (195, 198),      // 低温电流范围
            ["振动"] = (195, 200),      // 使用常温电流范围
            ["冲击"] = (195, 200),      // 使用常温电流范围
            // 以下试验后条件都使用常温电流范围
            ["低温试验后"] = (195, 200), // 使用常温电流范围
            ["高温试验后"] = (195, 200), // 使用常温电流范围
            ["功能振动试验后"] = (195, 200), // 使用常温电流范围
            ["功能冲击试验后"] = (195, 200)  // 使用常温电流范围
        };

        /// <summary>
        /// 输出功率范围
        /// </summary>
        public static (double Min, double Max) OutputPowerRange = (-1.5, 1.5); // dB

        /// <summary>
        /// 频率准确度范围
        /// </summary>
        public static (double Min, double Max) FrequencyAccuracyRange = (-0.2, 0.2); // ppm

        /// <summary>
        /// 相噪参数范围
        /// </summary>
        public static (double Min, double Max) PhaseNoise1KHzRange = (-97, -94); // dBc
        public static (double Min, double Max) PhaseNoise10KHzRange = (-104, -102); // dBc

        /// <summary>
        /// 杂散抑制范围
        /// </summary>
        public static (double Min, double Max) SpuriousSuppressionRange = (-80, -70); // dBc

        /// <summary>
        /// 谐波抑制范围
        /// </summary>
        public static (double Min, double Max) HarmonicSuppressionRange = (-66, -58); // dBc

        /// <summary>
        /// 频率切换时间范围
        /// </summary>
        public static (double Min, double Max) FrequencySwitchTimeRange = (45, 46); // μs

        /// <summary>
        /// 电压值（固定）
        /// </summary>
        public static double Voltage = 5.0; // V

        /// <summary>
        /// 上电冲击电流范围
        /// </summary>
        public static (double Min, double Max) PowerOnCurrentRange = (0.23, 0.25); // A

        /// <summary>
        /// 冲击电流持续时间范围
        /// </summary>
        public static (double Min, double Max) ImpactDurationRange = (0.05, 0.07); // ms

        // [用户可修改结束]
        // ============================================

        /// <summary>
        /// 关键频率点
        /// </summary>
        public static readonly int[] KeyFrequencies = { 1025, 1101, 1150 };

        /// <summary>
        /// 完整频率范围
        /// </summary>
        public static readonly (int Start, int End) FrequencyRange = (1025, 1150);

        /// <summary>
        /// 输出功率最大变化步长
        /// </summary>
        public const double MaxPowerChangeStep = 0.1; // dB

        /// <summary>
        /// 输出功率微小波动范围（用于添加真实感）
        /// </summary>
        public const double OutputPowerVariation = 0.2; // dB

        /// <summary>
        /// 频率分段规则（新规则）
        /// 格式: (起始频率, 结束频率) -> (使用的关键频率, 是否添加0.1dB)
        /// </summary>
        public static readonly Dictionary<(int Min, int Max), (int KeyFrequency, double Offset)> FrequencySegmentsNew =
            new Dictionary<(int, int), (int, double)>
            {
                [(1025, 1054)] = (1025, 0.0),    // 1025MHz-1054MHz 使用1025MHz的功率值
                [(1055, 1100)] = (1025, 0.1),    // 1055MHz-1100MHz 使用1025MHz的功率值 + 0.1dB
                [(1101, 1101)] = (1101, 0.0),    // 1101MHz 使用自身的功率值
                [(1102, 1127)] = (1101, 0.0),    // 1102MHz-1127MHz 使用1101MHz的功率值
                [(1128, 1150)] = (1150, 0.0)     // 1128MHz-1150MHz 使用1150MHz的功率值
            };

        /// <summary>
        /// 根据频率获取对应的关键频率点和偏移量
        /// </summary>
        public static (int KeyFrequency, double Offset) GetPowerSettingsForFrequency(int frequency)
        {
            foreach (var segment in FrequencySegmentsNew)
            {
                if (frequency >= segment.Key.Min && frequency <= segment.Key.Max)
                {
                    return segment.Value;
                }
            }

            // 默认返回1025MHz，无偏移
            Console.WriteLine($"警告: 频率 {frequency}MHz 不在任何分段中，使用1025MHz");
            return (1025, 0.0);
        }

        /// <summary>
        /// 根据频率获取对应的关键频率点
        /// </summary>
        public static int GetKeyFrequencyForFrequency(int frequency)
        {
            // 首先检查是否为关键频率点本身
            if (frequency == 1025 || frequency == 1101 || frequency == 1150)
            {
                return frequency;
            }

            // 然后检查分段规则
            //foreach (var segment in FrequencySegments)
            //{
            //    if (frequency >= segment.Key.Min && frequency <= segment.Key.Max)
            //    {
            //        return segment.Value;
            //    }
            //}

            // 默认返回1025
            Console.WriteLine($"警告: 频率 {frequency}MHz 不在任何分段中，使用1025MHz");
            return 1025;
        }
    }
}
