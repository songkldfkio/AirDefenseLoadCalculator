using System.Globalization;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace AirDefenseLoad.Controls;

/// <summary>
/// WinForms NumericUpDown 的 WPF 复刻版（白底、右侧箭头条、两位小数、超范围自动夹紧）。
/// 原程序里 K∑p / K∑q 的取值范围是 0.01 ~ 1.00，步长 0.01。
/// </summary>
public partial class WinNumericUpDown : UserControl
{
    public static readonly DependencyProperty ValueProperty = DependencyProperty.Register(
        nameof(Value), typeof(double), typeof(WinNumericUpDown),
        new FrameworkPropertyMetadata(0.9,
            FrameworkPropertyMetadataOptions.BindsTwoWayByDefault,
            OnValueChanged, CoerceValue));

    public WinNumericUpDown()
    {
        InitializeComponent();
        Loaded += (_, _) =>
        {
            CoerceValue(ValueProperty);
            UpdateText();
        };
    }

    /// <summary>最小值，默认 0.01。</summary>
    public double Minimum { get; set; } = 0.01;

    /// <summary>最大值，默认 1。</summary>
    public double Maximum { get; set; } = 1.0;

    /// <summary>每次增减的步长，默认 0.01。</summary>
    public double Increment { get; set; } = 0.01;

    /// <summary>小数位数，默认 2。</summary>
    public int Decimals { get; set; } = 2;

    public double Value
    {
        get => (double)GetValue(ValueProperty);
        set => SetValue(ValueProperty, value);
    }

    private static object CoerceValue(DependencyObject d, object baseValue)
    {
        var box = (WinNumericUpDown)d;
        var v = (double)baseValue;

        if (double.IsNaN(v)) v = box.Minimum;
        v = Math.Round(v, box.Decimals, MidpointRounding.AwayFromZero);

        if (v < box.Minimum) v = box.Minimum;
        if (v > box.Maximum) v = box.Maximum;
        return v;
    }

    private static void OnValueChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        => ((WinNumericUpDown)d).UpdateText();

    private void UpdateText()
    {
        if (Editor is null) return;

        var text = Value.ToString("0." + new string('0', Math.Max(0, Decimals)), CultureInfo.CurrentCulture);
        if (Editor.Text != text) Editor.Text = text;
    }

    private void Commit()
    {
        if (double.TryParse(Editor.Text.Trim(), NumberStyles.Float, CultureInfo.CurrentCulture, out var v))
            Value = v;
        else
            UpdateText();
    }

    private void Editor_LostFocus(object sender, RoutedEventArgs e) => Commit();

    private void Editor_KeyDown(object sender, KeyEventArgs e)
    {
        switch (e.Key)
        {
            case Key.Enter:
                Commit();
                e.Handled = true;
                break;

            case Key.Up:
                Value += Increment;
                e.Handled = true;
                break;

            case Key.Down:
                Value -= Increment;
                e.Handled = true;
                break;
        }
    }

    private void Editor_MouseWheel(object sender, MouseWheelEventArgs e)
    {
        Value += e.Delta > 0 ? Increment : -Increment;
        e.Handled = true;
    }

    private void Up_Click(object sender, RoutedEventArgs e) => Value += Increment;

    private void Down_Click(object sender, RoutedEventArgs e) => Value -= Increment;
}
