// Models/TestCondition.cs - 测试条件枚举
namespace TestReportAutoFiller.Models
{
    /// <summary>
    /// 测试条件类型
    /// </summary>
    public enum TestCondition
    {
        Normal,            // 常温
        LowTemp,           // 低温
        HighTemp,          // 高温
        Vibration,         // 功能振动
        Impact,            // 功能冲击
        AfterLowTemp,      // 低温试验后
        AfterHighTemp,     // 高温试验后
        AfterVibration,    // 功能振动试验后
        AfterImpact        // 功能冲击试验后
    }

    /// <summary>
    /// 测试环境参数
    /// </summary>
    public class TestEnvironment
    {
        public TestCondition Condition { get; set; }
        public string Temperature { get; set; }
        public string Humidity { get; set; }
        public DateTime TestTime { get; set; }
    }

    /// <summary>
    /// 产品信息
    /// </summary>
    public class ProductInfo
    {
        public string Name { get; set; } = "频率合成器";
        public string Model { get; set; } = "HD-FS-L011-1";
        public string SerialNumber { get; set; }
        public string TestCategory { get; set; } = "交付检验";
    }
}