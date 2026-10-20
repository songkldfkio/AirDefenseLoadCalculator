// .NET Framework 4.8 没有这两个类型，C# 9 的 init / record 需要它们才能编译。
// 仅作编译期标记，不参与运行。
namespace System.Runtime.CompilerServices
{
    internal static class IsExternalInit
    {
    }
}

namespace System
{
    /// <summary>net6+ 的 Math.Clamp 在 net48 里没有，补一个等价实现。</summary>
    internal static class MathCompat
    {
        public static double Clamp(double value, double min, double max)
        {
            if (value < min) return min;
            if (value > max) return max;
            return value;
        }

        public static int Clamp(int value, int min, int max)
        {
            if (value < min) return min;
            if (value > max) return max;
            return value;
        }
    }
}
