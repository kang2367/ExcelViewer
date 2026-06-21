using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;

namespace ExcelDropViewer
{
    internal static class ScrollableMessageDialog
    {
        public static void Show(Window? owner, string title, string message)
        {
            var maxWidth = SystemParameters.WorkArea.Width * 0.75;
            var maxHeight = SystemParameters.WorkArea.Height * 0.75;

            var messageBox = new TextBox
            {
                Text = message,
                IsReadOnly = true,
                AcceptsReturn = true,
                TextWrapping = TextWrapping.NoWrap,
                BorderThickness = new Thickness(0),
                Background = Brushes.Transparent,
                FontFamily = new FontFamily("Segoe UI"),
                FontSize = 14,
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                HorizontalScrollBarVisibility = ScrollBarVisibility.Auto,
                MaxWidth = maxWidth - 48,
                MaxHeight = maxHeight - 120
            };
            messageBox.PreviewMouseWheel += MessageBox_PreviewMouseWheel;

            var okButton = new Button
            {
                Content = "확인",
                IsDefault = true,
                MinWidth = 80,
                Margin = new Thickness(0, 8, 0, 0),
                HorizontalAlignment = HorizontalAlignment.Center,
                Padding = new Thickness(16, 6, 16, 6)
            };

            var root = new DockPanel
            {
                Margin = new Thickness(16),
                LastChildFill = true
            };

            DockPanel.SetDock(okButton, Dock.Bottom);
            root.Children.Add(okButton);
            root.Children.Add(messageBox);

            var window = new Window
            {
                Title = title,
                Owner = owner,
                Content = root,
                WindowStartupLocation = owner != null
                    ? WindowStartupLocation.CenterOwner
                    : WindowStartupLocation.CenterScreen,
                MinWidth = 320,
                MinHeight = 160,
                MaxWidth = maxWidth,
                MaxHeight = maxHeight,
                SizeToContent = SizeToContent.WidthAndHeight,
                ResizeMode = ResizeMode.CanResize,
                ShowInTaskbar = false
            };

            okButton.Click += (_, _) => window.Close();
            window.ShowDialog();
        }

        private static void MessageBox_PreviewMouseWheel(object sender, MouseWheelEventArgs e)
        {
            if (Keyboard.Modifiers != ModifierKeys.Shift || sender is not TextBox textBox)
            {
                return;
            }

            e.Handled = true;
            var nextOffset = textBox.HorizontalOffset - e.Delta;
            if (nextOffset < 0)
            {
                nextOffset = 0;
            }

            textBox.ScrollToHorizontalOffset(nextOffset);
        }
    }
}
