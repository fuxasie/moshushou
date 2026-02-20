using System.Windows;

namespace moshushou
{
    public partial class SettingsWindow : Window
    {
        private readonly SearchConfig _config;

        public SettingsWindow(SearchConfig config)
        {
            InitializeComponent();
            _config = config;

            // 初始化界面控件状态
            EnableOsdCheckBox.IsChecked = _config.EnableOsdWindow;
            SkipNextOnCtrlSpaceCheckBox.IsChecked = _config.SkipNextOnCtrlSpace;
        }

        private void EnableOsdCheckBox_Checked(object sender, RoutedEventArgs e)
        {
            _config.EnableOsdWindow = true;
            _config.Save();
        }

        private void EnableOsdCheckBox_Unchecked(object sender, RoutedEventArgs e)
        {
            _config.EnableOsdWindow = false;
            _config.Save();
        }

        private void OsdSettingsButton_Click(object sender, RoutedEventArgs e)
        {
            OsdWindow.ToggleEditMode();
        }

        private void SkipNextOnCtrlSpaceCheckBox_Checked(object sender, RoutedEventArgs e)
        {
            _config.SkipNextOnCtrlSpace = true;
            _config.Save();

            // 若需实时同步给 MainWindow，可在被调处读取 _config
        }

        private void SkipNextOnCtrlSpaceCheckBox_Unchecked(object sender, RoutedEventArgs e)
        {
            _config.SkipNextOnCtrlSpace = false;
            _config.Save();
        }

        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
    }
}
