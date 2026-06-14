using System;
using System.Globalization;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Threading;

namespace ExcelDropViewer
{
    internal sealed class UiLogWriter
    {
        private readonly TextBox _logTextBox;
        private readonly ScrollViewer? _scrollViewer;
        private readonly Dispatcher _dispatcher;

        public UiLogWriter(TextBox logTextBox, ScrollViewer? scrollViewer = null)
        {
            _logTextBox = logTextBox;
            _scrollViewer = scrollViewer;
            _dispatcher = logTextBox.Dispatcher;
        }

        public void LogStart(string functionName)
        {
            AppendLine($"{functionName} start.");
        }

        public void LogEnd(string functionName)
        {
            AppendLine($"{functionName} end.");
        }

        public void LogProgress(string functionName, string message)
        {
            AppendLine($"{functionName} {message}");
        }

        public void LogRowProgress(string functionName, int current, int total, string action)
        {
            AppendLine($"{functionName} {action} {current}/{total}.");
        }

        private void AppendLine(string message)
        {
            var line = $"[{DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.CurrentCulture)}]{message}";

            if (_dispatcher.CheckAccess())
            {
                WriteLine(line);
                return;
            }

            _dispatcher.Invoke(() => WriteLine(line));
        }

        private void WriteLine(string line)
        {
            if (string.IsNullOrEmpty(_logTextBox.Text))
            {
                _logTextBox.Text = line;
            }
            else
            {
                _logTextBox.AppendText(Environment.NewLine + line);
            }

            _logTextBox.CaretIndex = _logTextBox.Text.Length;
            _logTextBox.ScrollToEnd();

            if (_scrollViewer != null)
            {
                _scrollViewer.ScrollToVerticalOffset(_scrollViewer.ExtentHeight);
            }

            _logTextBox.Focus();
            Keyboard.Focus(_logTextBox);
        }
    }
}
