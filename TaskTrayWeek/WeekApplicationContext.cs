using System.Globalization;
using System.Runtime.InteropServices;
using Microsoft.Win32;

namespace TaskTrayWeek;

public class WeekApplicationContext : ApplicationContext
{
    private const int DefaultTrayIconSize = 16;
    private const int MinIconFontSize = 6;
    private const int SmCxSmIcon = 49;
    private const int SmCySmIcon = 50;
    private const string StartupRegistryPath = @"SOFTWARE\Microsoft\Windows\CurrentVersion\Run";
    private const string StartupRegistryName = "TaskTrayWeek";
    private readonly ToolStripMenuItem _exitApplicationItem;
    private readonly ToolStripTextBox _getWeekNumFromDateItem;
    private readonly ToolStripTextBox _getWeekNumItem;
    private readonly ToolStripMenuItem _manuallyUpdateIconItem;
    private readonly NotifyIcon _trayIcon;
    private readonly ToolStripMenuItem _toolTipStartupCheckbox;
    private readonly ToolStripMenuItem _toolTipTitleItem;
    private readonly System.Windows.Forms.Timer _timer;
    private bool _isAddedToStartup;

    public WeekApplicationContext()
    {
        _timer = CreateTimer();
        _isAddedToStartup = IsApplicationInStartup();
        _toolTipTitleItem = CreateTitleMenuItem();
        _manuallyUpdateIconItem = new ToolStripMenuItem("Force update", null, (_, _) => TimerOnElapsed(this, EventArgs.Empty));
        _exitApplicationItem = new ToolStripMenuItem("Exit", null, ExitApplication);
        _toolTipStartupCheckbox = CreateStartupMenuItem();
        _getWeekNumItem = CreateGetWeekNumItem();
        _getWeekNumFromDateItem = CreateGetWeekNumFromDateItem();
        _trayIcon = CreateTrayIcon();
        InitializeApplication();
    }

    private System.Windows.Forms.Timer CreateTimer()
    {
        var timer = new System.Windows.Forms.Timer
        {
            Interval = 3600000
        };
        timer.Tick += TimerOnElapsed;

        return timer;
    }

    private static ToolStripMenuItem CreateTitleMenuItem()
    {
        return new ToolStripMenuItem("Current week")
        {
            Enabled = false
        };
    }

    private ToolStripMenuItem CreateStartupMenuItem()
    {
        return new ToolStripMenuItem("Start on boot", null, AddToStartup)
        {
            Checked = _isAddedToStartup,
            CheckOnClick = false,
            ToolTipText = "Start TaskTrayWeek when you sign in to Windows."
        };
    }

    private ToolStripTextBox CreateGetWeekNumItem()
    {
        var item = new ToolStripTextBox("Get date from week")
        {
            ToolTipText = "Enter a week number, then press Enter to show the Monday of that week."
        };
        item.TextBox.PlaceholderText = "Enter weeknumber";
        item.KeyDown += GetWeekNumItemOnKeyDown;

        return item;
    }

    private ToolStripTextBox CreateGetWeekNumFromDateItem()
    {
        var item = new ToolStripTextBox("Get week from date")
        {
            ToolTipText = "Enter a date, then press Enter to show its week number."
        };
        item.TextBox.PlaceholderText = "Enter date";
        item.KeyDown += GetWeekNumFromDateItemOnKeyDown;

        return item;
    }

    private bool IsApplicationInStartup()
    {
        using (var key = Registry.CurrentUser.OpenSubKey(StartupRegistryPath, false))
        {
            return key?.GetValue(StartupRegistryName) != null;
        }
    }

    private void AddToStartup(object? sender, EventArgs e)
    {
        var shouldAddToStartup = !_isAddedToStartup;

        if (shouldAddToStartup)
        {
            AddApplicationToStartup();
        }
        else
        {
            RemoveApplicationFromStartup();
        }

        _isAddedToStartup = IsApplicationInStartup();
        _toolTipStartupCheckbox.Checked = _isAddedToStartup;
    }

    private static void AddApplicationToStartup()
    {
        try
        {
            using (var key = Registry.CurrentUser.OpenSubKey(StartupRegistryPath, true))
            {
                if (key == null)
                {
                    MessageBox.Show("Failed to open registry key.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                key.SetValue(StartupRegistryName, GetStartupCommand(), RegistryValueKind.String);
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Failed to add application to startup: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }


    private static void RemoveApplicationFromStartup()
    {
        try
        {
            using (var key = Registry.CurrentUser.OpenSubKey(StartupRegistryPath, true))
            {
                if (key == null)
                {
                    MessageBox.Show("Failed to open registry key.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                key.DeleteValue(StartupRegistryName, false);
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Failed to remove application from startup: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private static string GetStartupCommand()
    {
        var assemblyLocation = System.Reflection.Assembly.GetExecutingAssembly().Location;
        if (!string.IsNullOrWhiteSpace(assemblyLocation))
        {
            var appHostPath = Path.ChangeExtension(assemblyLocation, ".exe");
            if (File.Exists(appHostPath))
            {
                return QuoteCommandArgument(appHostPath);
            }

            if (File.Exists(assemblyLocation) && !string.IsNullOrWhiteSpace(Environment.ProcessPath))
            {
                return $"{QuoteCommandArgument(Environment.ProcessPath)} {QuoteCommandArgument(assemblyLocation)}";
            }
        }

        return QuoteCommandArgument(Application.ExecutablePath);
    }

    private static string QuoteCommandArgument(string value)
    {
        return $"\"{value}\"";
    }

    private NotifyIcon CreateTrayIcon()
    {
        return new NotifyIcon
        {
            Visible = true,
            Icon = GenerateIcon(0),
            BalloonTipText = "Current week",
            Text = "Current week",
            ContextMenuStrip = new ContextMenuStrip
            {
                ShowItemToolTips = true,
                Items =
                {
                    _toolTipTitleItem, _manuallyUpdateIconItem, _getWeekNumItem, _getWeekNumFromDateItem, _toolTipStartupCheckbox,
                    _exitApplicationItem
                }
            }
        };
    }

    private void InitializeApplication()
    {
        UpdateTrayIcon(GetWeekNumber(DateTime.Now));
        _timer.Start();
        Application.ApplicationExit += OnApplicationExit;
    }

    private void GetWeekNumItemOnKeyDown(object? sender, KeyEventArgs e)
    {
        if (e.KeyCode == Keys.Enter)
        {
            if (int.TryParse(_getWeekNumItem.TextBox.Text, out var result))
            {
                var weekNum = FirstDateOfWeekISO8601(result);
                _getWeekNumItem.TextBox.Text = weekNum.ToShortDateString();
            }
            else
            {
                _getWeekNumItem.TextBox.Text = "Badly formatted string";
            }
        }
    }

    private void GetWeekNumFromDateItemOnKeyDown(object? sender, KeyEventArgs e)
    {
        if (e.KeyCode == Keys.Enter)
        {
            if (DateOnly.TryParse(_getWeekNumFromDateItem.TextBox.Text, out var result))
            {
                var weekNum = GetWeekNumber(result);
                _getWeekNumFromDateItem.TextBox.Text = weekNum.ToString();
            }
            else
            {
                _getWeekNumFromDateItem.TextBox.Text = "Badly formatted string";
            }
        }
    }

    private void TimerOnElapsed(object? sender, EventArgs e)
    {
        UpdateTrayIcon(GetWeekNumber(DateTime.Now));
    }

    private static void ExitApplication(object? sender, EventArgs e)
    {
        Application.Exit();
    }

    private void OnApplicationExit(object? sender, EventArgs e)
    {
        _timer.Stop();
        _timer.Dispose();
        // Hide tray icon, otherwise it will remain shown until user mouses over it
        _trayIcon.Visible = false;
        _trayIcon.Dispose();
    }

    /// <summary>
    ///     Dynamically generate a weeknumber icon for the application
    /// </summary>
    /// <param name="weekNum">Weeknumber to generate icon for</param>
    /// <returns>A windows application icon</returns>
    private static Icon GenerateIcon(int weekNum)
    {
        var iconSize = GetScaledTrayIconSize();
        using var bitmap = new Bitmap(iconSize.Width, iconSize.Height);
        using (var g = Graphics.FromImage(bitmap))
        {
            g.Clear(Color.Black);

            var text = weekNum.ToString(CultureInfo.InvariantCulture);
            using var font = CreateIconFont(text, iconSize);
            TextRenderer.DrawText(g, text, font, new Rectangle(Point.Empty, iconSize), Color.White,
                TextFormatFlags.HorizontalCenter |
                TextFormatFlags.VerticalCenter |
                TextFormatFlags.SingleLine |
                TextFormatFlags.NoPadding |
                TextFormatFlags.NoPrefix);
        }

        var iconHandle = bitmap.GetHicon();
        try
        {
            using var icon = Icon.FromHandle(iconHandle);
            return (Icon)icon.Clone();
        }
        finally
        {
            DestroyIcon(iconHandle);
        }
    }

    private void UpdateTrayIcon(int weekNum)
    {
        var previousIcon = _trayIcon.Icon;
        _trayIcon.Icon = GenerateIcon(weekNum);
        previousIcon?.Dispose();
    }

    private static Size GetScaledTrayIconSize()
    {
        if (OperatingSystem.IsWindowsVersionAtLeast(10, 0, 14393))
        {
            var dpi = GetDpiForSystem();
            var width = GetSystemMetricsForDpi(SmCxSmIcon, dpi);
            var height = GetSystemMetricsForDpi(SmCySmIcon, dpi);

            if (width > 0 && height > 0)
            {
                return new Size(width, height);
            }
        }

        return SystemInformation.SmallIconSize is { Width: > 0, Height: > 0 } size
            ? size
            : new Size(DefaultTrayIconSize, DefaultTrayIconSize);
    }

    private static Font CreateIconFont(string text, Size iconSize)
    {
        for (var fontSize = iconSize.Height; fontSize >= MinIconFontSize; fontSize--)
        {
            var font = new Font("Verdana", fontSize, FontStyle.Bold, GraphicsUnit.Pixel);
            var measuredSize = TextRenderer.MeasureText(text, font, Size.Empty,
                TextFormatFlags.SingleLine | TextFormatFlags.NoPadding | TextFormatFlags.NoPrefix);

            if (measuredSize.Width <= iconSize.Width && measuredSize.Height <= iconSize.Height)
            {
                return font;
            }

            font.Dispose();
        }

        return new Font("Verdana", MinIconFontSize, FontStyle.Bold, GraphicsUnit.Pixel);
    }

    private int GetWeekNumber(DateTime date)
    {
        var ciCurr = CultureInfo.CurrentCulture;
        return ciCurr.Calendar.GetWeekOfYear(date, CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday);
    }

    private int GetWeekNumber(DateOnly date)
    {
        return GetWeekNumber(date.ToDateTime(TimeOnly.MinValue));
    }

    /// <summary>
    ///     Gets the date of the Monday in a week depending on the weeknumber following the ISO8601 standard
    /// </summary>
    /// <param name="weekOfYear">Week number to get the date from</param>
    /// <returns></returns>
    private DateTime FirstDateOfWeekISO8601(int weekOfYear)
    {
        var jan1 = new DateTime(DateTime.Now.Year, 1, 1);
        var daysOffset = DayOfWeek.Thursday - jan1.DayOfWeek;

        // Use first Thursday in January to get first week of the year as
        // it will never be in Week 52/53
        var firstThursday = jan1.AddDays(daysOffset);
        var cal = CultureInfo.CurrentCulture.Calendar;
        var firstWeek = cal.GetWeekOfYear(firstThursday, CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday);

        var weekNum = weekOfYear;
        // As we're adding days to a date in Week 1,
        // we need to subtract 1 in order to get the right date for week #1
        if (firstWeek == 1) weekNum -= 1;

        // Using the first Thursday as starting week ensures that we are starting in the right year
        // then we add number of weeks multiplied with days
        var result = firstThursday.AddDays(weekNum * 7);

        // Subtract 3 days from Thursday to get Monday, which is the first weekday in ISO8601
        return result.AddDays(-3);
    }

    [DllImport("user32.dll")]
    private static extern uint GetDpiForSystem();

    [DllImport("user32.dll")]
    private static extern int GetSystemMetricsForDpi(int nIndex, uint dpi);

    [DllImport("user32.dll")]
    private static extern bool DestroyIcon(IntPtr hIcon);
}
