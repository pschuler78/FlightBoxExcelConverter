using System;

namespace FlightBoxExcelConverter
{
    public class LogEventArgs : EventArgs
    {
        public string Text { get; set; }
        public LogEventArgs(string text)
        {
            Text = text;
        }
    }
}
