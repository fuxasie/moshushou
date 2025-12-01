using System.Windows;

namespace moshushou
{
    public partial class EditBusInfoWindow : Window
    {
        public BusinessInfo Info { get; private set; }

        public EditBusInfoWindow(BusinessInfo busInfo)
        {
            InitializeComponent();
            this.Owner = Application.Current.MainWindow; // Set owner for proper dialog behavior

            Info = busInfo;

            // Populate the UI with the existing data
            StoreNameTextBlock.Text = Info.StoreName;
            GroupNameTextBox.Text = Info.GroupName;

            if (!string.IsNullOrEmpty(Info.Source))
            {
                if (Info.Source.Equals("企业微信"))
                {
                    WeworkRadioButton.IsChecked = true;
                }
                else if (Info.Source.Equals("微信"))
                {
                    WechatRadioButton.IsChecked = true;
                }
            }
        }

        private void SaveButton_Click(object sender, RoutedEventArgs e)
        {
            // Update the Info object from the UI controls before closing
            Info.GroupName = GroupNameTextBox.Text.Trim();

            if (WeworkRadioButton.IsChecked == true)
            {
                Info.Source = "企业微信";
            }
            else if (WechatRadioButton.IsChecked == true)
            {
                Info.Source = "微信";
            }
            else
            {
                Info.Source = null; // No source selected
            }

            // If group name is empty, clear the source as well
            if (string.IsNullOrWhiteSpace(Info.GroupName))
            {
                Info.GroupName = null;
                Info.Source = null;
            }

            this.DialogResult = true;
            this.Close();
        }

        private void DeleteButton_Click(object sender, RoutedEventArgs e)
        {
                // Clear the data
                Info.GroupName = null;
                Info.Source = null;
                this.DialogResult = true;
                this.Close();
        }
    }
}