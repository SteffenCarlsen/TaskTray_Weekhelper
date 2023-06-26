using System.Globalization;
using System.Timers;
using Timer = System.Timers.Timer;

namespace TaskTrayWeek;

public class WeekApplicationContext : ApplicationContext
{
    private const int FONTSIZE_SMALL = 15;
    private const int FONTSIZE_BIG = 21;
    
    private readonly NotifyIcon _trayIcon;
    private readonly ToolStripTextBox _getWeekNUm; 
    public WeekApplicationContext ()
    {
        // Check if week has changed once every hour incase some pesky person forgot to shut their PC down
        var timer = new Timer();
        timer.Interval = 3600000;
        timer.Elapsed += TimerOnElapsed;
        
        var itemTop = new ToolStripMenuItem("Current week");
        itemTop.Enabled = false;
        
        var itemUpdate = new ToolStripMenuItem("Force update", null, delegate { TimerOnElapsed(null ,null!); });
        
        _getWeekNUm = new ToolStripTextBox("Get date from week");
        _getWeekNUm.TextBox!.PlaceholderText = "Enter week number and press enter";
        _getWeekNUm.KeyDown += GetWeekNUmOnKeyDown;
        
        var item = new ToolStripMenuItem("Exit", null, OnApplicationExit);
        
        // Initialize Tray Icon
        _trayIcon = new NotifyIcon();
        _trayIcon.Visible = true;
        _trayIcon.ContextMenuStrip = new ContextMenuStrip() { Items = { itemTop, itemUpdate, _getWeekNUm, item }};
        _trayIcon.Icon = GenerateIcon(0);
        _trayIcon.BalloonTipText = "Current week";
        _trayIcon.Text = "Current week";
        
        // Initialize application
        TimerOnElapsed(null,null!);
        Application.ApplicationExit += OnApplicationExit;
        timer.Start();
    }

    private void GetWeekNUmOnKeyDown(object? sender, KeyEventArgs e)
    {
        if (e.KeyCode == Keys.Enter && _getWeekNUm.TextBox != null)
        {
            if (int.TryParse(_getWeekNUm.TextBox.Text, out var result))
            {
                var weekNum = FirstDateOfWeekISO8601(result);
                _getWeekNUm.TextBox.Text = weekNum.ToShortDateString();
            }
            else
            {
                _getWeekNUm.TextBox.Text = "Badly formatted string";
            }
        }
    }

    private void TimerOnElapsed(object? sender, ElapsedEventArgs e)
    {
        _trayIcon.Icon = GenerateIcon(GetWeekNumber(DateTime.Now));
    }

    private void OnApplicationExit(object? sender, EventArgs e)
    {
        // Hide tray icon, otherwise it will remain shown until user mouses over it
        _trayIcon.Visible = false;
        Application.Exit();
    }

    /// <summary>
    /// Dynamically generate a weeknumber icon for the application
    /// </summary>
    /// <param name="weekNum">Weeknumber to generate icon for</param>
    /// <returns>A windows application icon</returns>
    private Icon GenerateIcon(int weekNum)
    {
        var bitmap = new Bitmap(32, 32);
        using (var g = Graphics.FromImage(bitmap))
        {
            g.FillRectangle(new SolidBrush(Color.Black), new Rectangle(0,0,32,32));
            g.DrawString(weekNum.ToString(), new Font("Verdana", weekNum > 9 ? FONTSIZE_SMALL : FONTSIZE_BIG, FontStyle.Bold), new SolidBrush(Color.White), new PointF(0, 2));
            return Icon.FromHandle(bitmap.GetHicon());
        }
        
    }

    private int GetWeekNumber(DateTime date)
    {
        var ciCurr = CultureInfo.CurrentCulture;
        return ciCurr.Calendar.GetWeekOfYear(date, CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday);
    }
    
    /// <summary>
    /// Gets the date of the Monday in a week depending on the weeknumber following the ISO8601 standard
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
        if (firstWeek == 1)
        {
            weekNum -= 1;
        }

        // Using the first Thursday as starting week ensures that we are starting in the right year
        // then we add number of weeks multiplied with days
        var result = firstThursday.AddDays(weekNum * 7);

        // Subtract 3 days from Thursday to get Monday, which is the first weekday in ISO8601
        return result.AddDays(-3);
    }    
    
}