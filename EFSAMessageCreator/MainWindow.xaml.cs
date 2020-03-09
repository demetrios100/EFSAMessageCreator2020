// <copyright file="MainWindow.xaml.cs" company="EFSAUsersGroup">Copyright (c) EFSA Users Group. All rights reserved.</copyright>
// <author>Demetrios Ioannides</author>
// <email>dvi1@columbia.edu</email>
// <summary>The Main Window for the EFSA Message Creator Utility</summary>

namespace EFSAMessageCreator
{
    #region using statements
    using System;
    using System.Configuration;
    using System.IO;
    using System.Linq;
    using System.Windows;
    using System.Windows.Controls;
    using System.Windows.Media;
    using System.Windows.Media.Imaging;
    #endregion

    /// <summary>
    /// File Type
    /// </summary>
    public enum FileType
    {
        /// <summary>
        /// DBF File
        /// </summary>
        DBF,

        /// <summary>
        /// Delimiter File
        /// </summary>
        Delimited
    }

    /// <summary>
    /// Delimiter Type
    /// </summary>
    public enum DelimiterType
    {
        /// <summary>
        /// Delimited with Comma with Quotes
        /// </summary>
        CommaWithQuotes,

        /// <summary>
        /// Delimited with Semicolon
        /// </summary>
        Semicolon,

        /// <summary>
        /// Delimited with Tab
        /// </summary>
        Tab,

        /// <summary>
        /// Delimited with Pipe (Vertical Bar)
        /// </summary>
        Pipe
    }

    /// <summary>
    /// Interaction logic for the Main Window
    /// </summary>
    public partial class MainWindow : Window
    {
        /// <summary>
        /// The Encoding of the Output XML file
        /// </summary>
        private System.Text.Encoding outputXMLEncoding = System.Text.Encoding.Unicode;

        /// <summary>
        /// The Type of the Data File (Delimited or DBF)
        /// </summary>
        private FileType fileType;

        /// <summary>
        /// The Delimiter Type
        /// </summary>
        private DelimiterType delimiterType = DelimiterType.CommaWithQuotes;

        #region Initialization
        /// <summary>
        /// Initializes a new instance of the <see cref="MainWindow"/> class
        /// </summary>
        public MainWindow()
        {
            this.InitializeComponent();

            this.Title = App.ApplicationTitle;
            AccessText text = new AccessText();
            text.TextWrapping = TextWrapping.Wrap;
            text.TextAlignment = TextAlignment.Center;
            text.Text = App.ApplicationTitle;
            appBanner.Content = text;
            Uri iconUri = new Uri("pack://application:,,,/" + App.ApplicationIcon, UriKind.RelativeOrAbsolute);
            this.Icon = BitmapFrame.Create(iconUri);
            var bc = new BrushConverter();
            this.Background = (Brush)bc.ConvertFrom(App.ApplicationBackground);
            SchemaLabel.Content = "Schema: " + App.Schema;

            if (ConfigurationManager.AppSettings["datapath"] != null)
            {
                DataFileTextBox.Text = ConfigurationManager.AppSettings["datapath"];
            }

            ComboBoxItem cbi;
            cbi = new ComboBoxItem
            {
                Name = "CommaWithQuotes",
                Content = "Comma ,"
            };
            DelimiterCheckBox.Items.Add(cbi);

            cbi = new ComboBoxItem
            {
                Name = "Semicolon",
                Content = "Semicolon ;"
            };
            DelimiterCheckBox.Items.Add(cbi);

            cbi = new ComboBoxItem
            {
                Name = "Tab",
                Content = "Tab"
            };
            DelimiterCheckBox.Items.Add(cbi);

            cbi = new ComboBoxItem
            {
                Name = "Pipe",
                Content = "Pipe (Vertical Bar) |"
            };

            DelimiterCheckBox.Items.Add(cbi);
        }
        #endregion

        #region Select the Data File
        /// <summary>
        /// Select the Data File
        /// </summary>
        /// <param name="sender">The Object sender</param>
        /// <param name="e"> The Routed arguments</param>
        private void SelectButton_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new System.Windows.Forms.OpenFileDialog();
            dialog.Filter = "Text files(*.txt) |*.txt| CSV Text files(*.csv) |*.csv| DBF files(*.dbf)|*.dbf| All files(*.*) | *.*";
            System.Windows.Forms.DialogResult result = dialog.ShowDialog();
            DataFileTextBox.Text = dialog.FileName;

            if (ConfigurationManager.AppSettings["datapath"] == null)
            {
                Configuration config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
                config.AppSettings.Settings.Add("datapath", dialog.FileName);
                config.Save(ConfigurationSaveMode.Full);
                ConfigurationManager.RefreshSection("appsettings");
            }
            else
            {
                Configuration config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
                config.AppSettings.Settings["datapath"].Value = dialog.FileName;
                config.Save(ConfigurationSaveMode.Full);
                ConfigurationManager.RefreshSection("appsettings");
            }
        }
        #endregion

        #region Button creating XML
        /// <summary>
        /// Button creating XML
        /// </summary>
        /// <param name="sender">The Object sender</param>
        /// <param name="e">The Routed Event arguments</param>
        private void CreateXMLButton_Click(object sender, RoutedEventArgs e)
        {
                try
                {
                    if (this.CheckFileType() == true)
                    {
                        this.CreateXML();
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
        }
        #endregion

        #region Create XML
        /// <summary>
        /// Create XML
        /// </summary>
        private void CreateXML()
        {
            string xsdFileFullName = "EFSAMessageCreator." + App.Schema.Trim();
            string elementMappingFileFullName = "EFSAMessageCreator." + App.ElementMappingFileName.Trim();

            XMLGeneration xmlGeneration
                = new XMLGeneration(
                    this,
                    xsdFileFullName,
                    elementMappingFileFullName,
                    App.OutputXMLFileName,
                    this.outputXMLEncoding,
                    this.fileType,
                    this.delimiterType,
                    DataFileTextBox.Text.Trim());
        }
        #endregion

        #region Limit CheckBox Handlers
        /// <summary>
        /// Limit CheckBox Checked Event Handler
        /// </summary>
        /// <param name="sender">The Object Sender</param>
        /// <param name="e">The Routed Event Arguments</param>
        private void LimitCheckBox_Checked(object sender, RoutedEventArgs e)
        {
            LimitTextBox.Visibility = Visibility.Visible;
            LimitCheckBox.Content = "For testing, limit the number of records generated to ";
        }

        /// <summary>
        /// Limit CheckBox Unchecked Event Handler
        /// </summary>
        /// <param name="sender">The Object Sender</param>
        /// <param name="e">The Routed Event Arguments</param>
        private void LimitCheckBox_Unchecked(object sender, RoutedEventArgs e)
        {
            LimitTextBox.Visibility = Visibility.Collapsed;
            LimitTextBox.Text = string.Empty;
            LimitCheckBox.Content = "For testing, limit the number of records generated";
        }
        #endregion

        #region Closing
        /// <summary>
        /// Window Closed
        /// </summary>
        /// <param name="sender">The Sender Object </param>
        /// <param name="e">The event Arguments</param>
        private void Window_Closed(object sender, EventArgs e)
        {
            // Environment.Exit(0);
            Application.Current.Shutdown();
        }
        #endregion

        #region Delimited Options
        /// <summary>
        /// The TextChanged event handler of DataFileTextBox
        /// </summary>
        /// <param name="sender">The sender object</param>
        /// <param name="e">The Arguments</param>
        private void DataFileTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            this.DetermineCSVOptionsDisplay();
        }

        /// <summary>
        /// The DataFileTextBox loaded event handler
        /// </summary>
        /// <param name="sender">The Sender Object</param>
        /// <param name="e">The Arguments</param>
        private void DataFileTextBox_Loaded(object sender, RoutedEventArgs e)
        {
            this.DetermineCSVOptionsDisplay();
        }

        /// <summary>
        /// Determine whether the CSV Options are being Displayed
        /// </summary>
        private void DetermineCSVOptionsDisplay()
        {
            string[] csvExtensions = { ".csv", ".txt" };

            if (DataFileTextBox.Text != string.Empty)
            {
                try
                {
                    FileInfo fileInfoDataFile = new FileInfo(DataFileTextBox.Text);
                    if (fileInfoDataFile.Extension.ToLower() == ".dbf")
                    {
                        pnlDelimitedOptions.Visibility = Visibility.Collapsed;
                    }
                    else if (csvExtensions.Contains(fileInfoDataFile.Extension.ToLower()))
                    {
                        pnlDelimitedOptions.Visibility = Visibility.Visible;
                    }
                }
                catch
                {
                    pnlDelimitedOptions.Visibility = Visibility.Collapsed;
                }
            }
        }
        #endregion

        #region Determine the File Type and ensure that if the file is delimited the delimiter has been specified
        /// <summary>
        /// Determine the File Type and ensure that if the file is delimited the delimiter has been specified
        /// </summary>
        /// <returns>A boolean indicating whether the file was selected successfully</returns>
        private bool CheckFileType()
        {
            bool fileSelectionOK = true;

            FileInfo fileInfoDataFile = new FileInfo(DataFileTextBox.Text.Trim());

            string[] csvExtensions = { ".csv", ".txt" };

            if (fileInfoDataFile.Extension.ToLower() == ".dbf")
            {
                this.fileType = FileType.DBF;
            }
            else if (csvExtensions.Contains(fileInfoDataFile.Extension.ToLower()))
            {
                this.fileType = FileType.Delimited;

                this.delimiterType = this.GetDelimiterType();

                if (DelimiterCheckBox.SelectedIndex == -1)
                {
                    fileSelectionOK = false;
                    MessageBox.Show("Please select a Delimiter", "Delimiter Missing");
                }
            }
            else
            {
                fileSelectionOK = false;
                MessageBox.Show("Invalid file extension");
            }

            return fileSelectionOK;
        }
        #endregion

        /// <summary>
        /// Get the Delimiter Type
        /// </summary>
        /// <returns>The delimiter Type</returns>
        private DelimiterType GetDelimiterType()
        {
            ComboBoxItem selectedDelimiter = (ComboBoxItem)DelimiterCheckBox.SelectedValue;

            switch (selectedDelimiter.Name)
            {
                case "CommaWithQuotes":
                    return DelimiterType.CommaWithQuotes;
                case "Pipe":
                    return DelimiterType.Pipe;
                case "Semicolon":
                    return DelimiterType.Semicolon;
                case "Tab":
                    return DelimiterType.Tab;
                default:
                    return DelimiterType.CommaWithQuotes;
            }
        }
    }
}
