// Models/TestConditionWithIndex.cs - 带索引的测试条件
namespace TestReportAutoFiller.Models
{
    /// <summary>
    /// 带索引的测试条件（用于筛选报告，确保每个表格的关键频点功率值独立）
    /// </summary>
    public class TestConditionWithIndex
    {
        public TestCondition Condition { get; }
        public int TableIndex { get; }

        public TestConditionWithIndex(TestCondition condition, int tableIndex)
        {
            Condition = condition;
            TableIndex = tableIndex;
        }

        public override bool Equals(object? obj)
        {
            if (obj is TestConditionWithIndex other)
            {
                return Condition == other.Condition && TableIndex == other.TableIndex;
            }
            return false;
        }

        public override int GetHashCode()
        {
            return HashCode.Combine(Condition, TableIndex);
        }
    }
}
