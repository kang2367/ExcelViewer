using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace ExcelDropViewer
{
    internal sealed class BomDuplicateItemDialog : Window
    {
        public BomDuplicateResolution Resolution { get; private set; } = BomDuplicateResolution.Cancel;

        private BomDuplicateItemDialog(string itemNumber)
        {
            Title = "품목 번호 중복";
            Width = 520;
            SizeToContent = SizeToContent.Height;
            WindowStartupLocation = WindowStartupLocation.CenterOwner;
            ResizeMode = ResizeMode.NoResize;
            ShowInTaskbar = false;

            var message = new TextBlock
            {
                Text = $"이미 존재 하는 품목 번호입니다: {itemNumber}\n기존 데이터를 새로운 정보로 업데이트하시겠습니까?",
                TextWrapping = TextWrapping.Wrap,
                Margin = new Thickness(0, 0, 0, 16)
            };

            var primaryButtonPanel = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Right,
                Margin = new Thickness(0, 0, 0, 8)
            };
            primaryButtonPanel.Children.Add(CreateButton("Update(업데이트)", () => CloseWith(BomDuplicateResolution.Update)));
            primaryButtonPanel.Children.Add(CreateButton("Skip(건너뛰기)", () => CloseWith(BomDuplicateResolution.Skip)));

            var secondaryButtonPanel = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Right
            };
            secondaryButtonPanel.Children.Add(CreateButton("All Skip(모두 건너뛰기)", () => CloseWith(BomDuplicateResolution.AllSkip)));
            secondaryButtonPanel.Children.Add(CreateButton("Cancel(취소)", () => CloseWith(BomDuplicateResolution.Cancel)));

            Content = new StackPanel
            {
                Margin = new Thickness(20),
                Children =
                {
                    message,
                    primaryButtonPanel,
                    secondaryButtonPanel
                }
            };
        }

        public static BomDuplicateResolution Show(Window owner, string itemNumber)
        {
            var previousCursor = Mouse.OverrideCursor;
            Mouse.OverrideCursor = null;

            try
            {
                var dialog = new BomDuplicateItemDialog(itemNumber)
                {
                    Owner = owner
                };
                dialog.ShowDialog();
                return dialog.Resolution;
            }
            finally
            {
                Mouse.OverrideCursor = previousCursor;
            }
        }

        private void CloseWith(BomDuplicateResolution resolution)
        {
            Resolution = resolution;
            DialogResult = resolution != BomDuplicateResolution.Cancel;
            Close();
        }

        private static System.Windows.Controls.Button CreateButton(string content, System.Action onClick)
        {
            var button = new System.Windows.Controls.Button
            {
                Content = content,
                MinWidth = 120,
                Margin = new Thickness(6, 0, 0, 0),
                Padding = new Thickness(12, 6, 12, 6)
            };
            button.Click += (_, _) => onClick();
            return button;
        }
    }
}
