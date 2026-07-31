using System.Windows;
using System.Windows.Controls;

namespace ExcelDropViewer
{
    public partial class DigikeyConfigWindow : Window
    {
        private string _savedClientId = string.Empty;
        private string _savedClientSecret = string.Empty;
        private bool _isInitializing;

        public DigikeyConfigWindow()
        {
            InitializeComponent();
            Loaded += DigikeyConfigWindow_Loaded;
        }

        private void DigikeyConfigWindow_Loaded(object sender, RoutedEventArgs e)
        {
            _isInitializing = true;
            try
            {
                var config = DigikeyConfigStore.Load();
                ClientIdTextBox.Text = config.ClientId;
                ClientSecretPasswordBox.Password = config.ClientSecret;
                RememberSavedState();
                UpdateApplyButtonState();
            }
            finally
            {
                _isInitializing = false;
            }
        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            SaveCurrentValues();
            DialogResult = true;
            Close();
        }

        private void ApplyButton_Click(object sender, RoutedEventArgs e)
        {
            SaveCurrentValues();
            ShowStatusMessage("설정이 저장되었습니다.");
            UpdateApplyButtonState();
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }

        private void Input_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (_isInitializing)
            {
                return;
            }

            HideStatusMessage();
            UpdateApplyButtonState();
        }

        private void Input_PasswordChanged(object sender, RoutedEventArgs e)
        {
            if (_isInitializing)
            {
                return;
            }

            HideStatusMessage();
            UpdateApplyButtonState();
        }

        private void SaveCurrentValues()
        {
            DigikeyConfigStore.Save(new DigikeyConfig
            {
                ClientId = ClientIdTextBox.Text.Trim(),
                ClientSecret = ClientSecretPasswordBox.Password
            });
            RememberSavedState();
        }

        private void RememberSavedState()
        {
            _savedClientId = ClientIdTextBox.Text.Trim();
            _savedClientSecret = ClientSecretPasswordBox.Password;
        }

        private void UpdateApplyButtonState()
        {
            var hasChanges = !string.Equals(ClientIdTextBox.Text.Trim(), _savedClientId, StringComparison.Ordinal)
                || !string.Equals(ClientSecretPasswordBox.Password, _savedClientSecret, StringComparison.Ordinal);
            ApplyButton.IsEnabled = hasChanges;
        }

        private void ShowStatusMessage(string message)
        {
            StatusTextBlock.Text = message;
            StatusTextBlock.Visibility = Visibility.Visible;
        }

        private void HideStatusMessage()
        {
            StatusTextBlock.Text = string.Empty;
            StatusTextBlock.Visibility = Visibility.Collapsed;
        }
    }
}
