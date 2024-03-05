using System.Globalization;
using Timer = System.Timers.Timer;

namespace TaskTrayWeek;

public class WeekApplicationContext : ApplicationContext
{
    private const int FONTSIZE_SMALL = 15;
    private const int FONTSIZE_BIG = 21;
    private ToolStripMenuItem _exitApplicationItem;
    private ToolStripTextBox _getWeekNumFromDateItem;
    private ToolStripTextBox _getWeekNumItem;
    private ToolStripMenuItem _manuallyUpdateIconItem;
    private ToolStripMenuItem _toolTipTitleItem;

    private NotifyIcon _trayIcon;

    public WeekApplicationContext()
    {
        InitializeTimer();
        InitializeMenuItem();
        InitializeToolStripTextBoxes();
        InitializeTrayIcon();
        InitializeApplication();
    }

    private void InitializeTimer()
    {
        var timer = new Timer
        {
            Interval = 3600000
        };
        timer.Elapsed += TimerOnElapsed;
        timer.Start();
    }

    private void InitializeMenuItem()
    {
        _toolTipTitleItem = new ToolStripMenuItem("Current week")
        {
            Enabled = false
        };
        _manuallyUpdateIconItem = new ToolStripMenuItem("Force update", null, (_, _) => TimerOnElapsed(this, EventArgs.Empty));
        _exitApplicationItem = new ToolStripMenuItem("Exit", null, OnApplicationExit);
    }

    private void InitializeToolStripTextBoxes()
    {
        _getWeekNumItem = new ToolStripTextBox("Get date from week");
        _getWeekNumItem.TextBox!.PlaceholderText = "Enter weeknumber";
        _getWeekNumItem.KeyDown += GetWeekNumItemOnKeyDown;

        _getWeekNumFromDateItem = new ToolStripTextBox("Get week from date");
        _getWeekNumFromDateItem.TextBox!.PlaceholderText = "Enter date";
        _getWeekNumFromDateItem.KeyDown += GetWeekNumFromDateItemOnKeyDown;
    }

    private void InitializeTrayIcon()
    {
        _trayIcon = new NotifyIcon
        {
            Visible = true,
            Icon = GenerateIcon(0),
            BalloonTipText = "Current week",
            Text = "Current week",
            ContextMenuStrip = new ContextMenuStrip
            {
                Items =
                {
                    _toolTipTitleItem, _manuallyUpdateIconItem, _getWeekNumItem, _getWeekNumFromDateItem,
                    _exitApplicationItem
                }
            }
        };
    }

    private void InitializeApplication()
    {
        TimerOnElapsed(this, EventArgs.Empty);
        Application.ApplicationExit += OnApplicationExit;
    }

    private void GetWeekNumItemOnKeyDown(object? sender, KeyEventArgs e)
    {
        if (e.KeyCode == Keys.Enter && _getWeekNumItem.TextBox != null)
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
        if (e.KeyCode == Keys.Enter && _getWeekNumFromDateItem.TextBox != null)
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
        _trayIcon.Icon = GenerateIcon(GetWeekNumber(DateTime.Now));
    }

    private void OnApplicationExit(object? sender, EventArgs e)
    {
        // Hide tray icon, otherwise it will remain shown until user mouses over it
        _trayIcon.Visible = false;
        _trayIcon.Dispose();
        Application.Exit();
    }

    /// <summary>
    ///     Dynamically generate a weeknumber icon for the application
    /// </summary>
    /// <param name="weekNum">Weeknumber to generate icon for</param>
    /// <returns>A windows application icon</returns>
    private Icon GenerateIcon(int weekNum)
    {
        var bitmap = new Bitmap(32, 32);
        using (var g = Graphics.FromImage(bitmap))
        {
            g.FillRectangle(new SolidBrush(Color.Black), new Rectangle(0, 0, 32, 32));
            g.DrawString(weekNum.ToString(),
                new Font("Verdana", weekNum > 9 ? FONTSIZE_SMALL : FONTSIZE_BIG, FontStyle.Bold),
                new SolidBrush(Color.White), new PointF(0, 2));
            return Icon.FromHandle(bitmap.GetHicon());
        }
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
}