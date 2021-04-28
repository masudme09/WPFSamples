using PDFSecureWPF.Classes;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;



namespace PDFSecureWPF
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        SolidColorBrush LightGrayBrush = new SolidColorBrush(Color.FromRgb(211, 211, 211));
        SolidColorBrush SelectedBrush = new SolidColorBrush(Color.FromRgb(255, 0, 0));

        public static string passwordInputPdfSecure="";
        public static string passwordInputBulkSecure="";
        public static int selectedOutputQuality;
        public static List<string> selectedFilePath;
        public double PdfSecureProgressBarValue;
        public SecurePdf securePdf;
        BackgroundWorker backWorker;
        bool isBatch;
        public static string currentProcessingFile;

        public MainWindow()
        {
            InitializeComponent();
            AllDefault();
            GridPdfSecure.Visibility = Visibility.Visible;
            lblSelectionPDFSecure.Background = SelectedBrush;

            #region screeSizeControl
            double width = System.Windows.SystemParameters.WorkArea.Width;
            double height = System.Windows.SystemParameters.WorkArea.Height;

            if(height<750)
            {
                MainWindowPDFSecure.Width = 700;
                MainWindowPDFSecure.Height = 600;
                //setDefaultLabeSize(10);
            }else
            {
                MainWindowPDFSecure.Height = 740;
                MainWindowPDFSecure.Width = 900;
            }

            #endregion screeSizeControl

            backWorker = new BackgroundWorker();
            backWorker.WorkerReportsProgress = true;
            backWorker.WorkerSupportsCancellation = true;
            backWorker.DoWork += BackWorker_DoWork;
            backWorker.ProgressChanged += BackWorker_ProgressChanged;
            selectedFilePath = new List<string>();
            isBatch = false;

        }

        private void BackWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            ProgressBar prBar;

            if (isBatch == true)
            {
                prBar = batchSecureProgressBar;
                lblFileProcessing.Content = currentProcessingFile;
            }
            else
                prBar = makeSecureProgressBar;

            prBar.Visibility = System.Windows.Visibility.Visible;
            if ((float)e.ProgressPercentage / 100.0 * prBar.Maximum > prBar.Maximum)
            {
                //prBar.Value = prBar.Maximum;
            }
            else if ((float)e.ProgressPercentage / 100.0 * prBar.Maximum == prBar.Maximum)
            {
                prBar.Visibility = System.Windows.Visibility.Hidden;
            }
            else
            {
                prBar.Value = (int)((float)e.ProgressPercentage / 100.0 * prBar.Maximum);
            }
        }

        private void BackWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            BackgroundWorker worker = sender as BackgroundWorker;
            securePdf.makeSecure(worker); //to secure the pdf
        }

        private void Label_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton == MouseButton.Left)
                Application.Current.MainWindow.DragMove();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        
        private void lblPdfSecure_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton == MouseButton.Left)
            {
                AllDefault();
                lblSelectionPDFSecure.Background = SelectedBrush;
                GridPdfSecure.Visibility = Visibility.Visible;
               
            }

        }

        private void lblBulkSecure_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton == MouseButton.Left)
            {
                AllDefault();
                lblSelectionBulkSecure.Background = SelectedBrush;
                GridBulkSecure.Visibility = Visibility.Visible;
            }
        }

        private void lblHelp_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton == MouseButton.Left)
            {
                AllDefault();
                lblSelectionHelp.Background = SelectedBrush;
                gridHelp.Visibility = Visibility.Visible;
            }
        }

        private void lblAbout_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton == MouseButton.Left)
            {
                AllDefault();
                lblSelectionAbout.Background = SelectedBrush;
                gridAbout.Visibility = Visibility.Visible;
                
            }
        }

        public void AllDefault()
        {
            GridPdfSecure.Visibility = Visibility.Hidden;
            gridAbout.Visibility = Visibility.Hidden;
            GridBulkSecure.Visibility = Visibility.Hidden;
            gridHelp.Visibility = Visibility.Hidden;
            lblSelectionAbout.Background = LightGrayBrush;
            lblSelectionBulkSecure.Background = LightGrayBrush;
            lblSelectionHelp.Background = LightGrayBrush;
            lblSelectionPDFSecure.Background = LightGrayBrush;
        }

        private void CheckBox_Checked(object sender, RoutedEventArgs e)
        {

        }

        private void chkBxPdfSecureZipPass_Checked(object sender, RoutedEventArgs e)
        {
            txtBoxPdfSecureZipPass.IsEnabled = true;
            Password.IsEnabled = true;
        }

        private void txtBoxPdfSecureZipPass_MouseEnter(object sender, MouseEventArgs e)
        {
            if (txtBoxPdfSecureZipPass.Text == "Password for ZIP")
                txtBoxPdfSecureZipPass.Text = "";
        }

        private void chkBxPdfSecureZipPass_Unchecked(object sender, RoutedEventArgs e)
        {            
            txtBoxPdfSecureZipPass.Text = "Password for ZIP";
            txtBoxPdfSecureZipPass.IsEnabled = false;
            Password.IsEnabled = false;
        }

        private void chkBulkSecureZipPass_Checked(object sender, RoutedEventArgs e)
        {
            txtBxBulkSecureZipPass.IsEnabled = true;
        }

        private void chkBulkSecureZipPass_Unchecked(object sender, RoutedEventArgs e)
        {
            txtBxBulkSecureZipPass.IsEnabled = false;
            txtBxBulkSecureZipPass.Text = "Password for ZIP";
        }

        private void txtBxBulkSecureZipPass_MouseEnter(object sender, MouseEventArgs e)
        {
            if(txtBxBulkSecureZipPass.Text== "Password for ZIP")
                txtBxBulkSecureZipPass.Text = "";
        }


        //Setting default fontsize of a Label
        public void setDefaultLabeSize(double lblFontSize)
        {
            Style lblStyle = this.FindResource("defaultLblStyle") as Style;
            foreach (Setter s in lblStyle.Setters)
            {
                if (s.Property == Label.FontSizeProperty)
                {
                    s.Value = lblFontSize;
                }
            }
        }

        private void btnSelectPDF_Click(object sender, RoutedEventArgs e)
        {
            

            string browsedPath = Utility.Browse();
            if (browsedPath != null)
            {
                selectedFilePath.Clear();//CLear all elements before adding new one
                txtBxFilePath.Text = browsedPath;
                selectedFilePath.Add(browsedPath);
            }
        }       

        private void radioButton96_Checked(object sender, RoutedEventArgs e)
        {
            RadioButton radio = sender as RadioButton;
            switch(radio.Content.ToString())
            {
                case "96":
                    selectedOutputQuality = 96;
                    break;
                case "240":
                    selectedOutputQuality = 240;
                    break;
                case "300":
                    selectedOutputQuality = 300;
                    break;
                case "600":
                    selectedOutputQuality = 600;
                    break;
                case "1200":
                    selectedOutputQuality = 1200;
                    break;                
            }
                

        }

        private void btnMakeSecurePDF_Click(object sender, RoutedEventArgs e)
        {
            securePdf = new SecurePdf(selectedOutputQuality, selectedFilePath, passwordInputPdfSecure);
            if(backWorker.IsBusy==false)
                backWorker.RunWorkerAsync();
        }

        private void Password_PasswordChanged(object sender, RoutedEventArgs e)
        {
            passwordInputPdfSecure = Password.Password;
        }


        //To update progressBarValue
        public void UpdateProgressBar(int value)
        {
            if (CheckAccess())
                this.makeSecureProgressBar.Value = value;
            else
            {
                Dispatcher.Invoke(new Action(() => this.makeSecureProgressBar.Value = value), null);
            }
        }

        private void btnSelectPdfFiles_Click(object sender, RoutedEventArgs e)
        {
            string[] files = Utility.BrowsePdfFiles();
            listBoxFiles.ItemsSource = files;
            selectedFilePath = files.ToList();
        }

        private void pdfBatchSecuring_Click(object sender, RoutedEventArgs e)
        {
            string saveFilePath="";
            isBatch = true;
            bool convertToSingle=false;
            if (chkBxConvertToSingle.IsChecked == true)
                convertToSingle = true;

            if (convertToSingle == true)
                saveFilePath = Utility.SaveFile();

            securePdf = new SecurePdf(300, selectedFilePath, passwordInputBulkSecure, convertToSingle,saveFilePath);
             if(backWorker.IsBusy==false)
                backWorker.RunWorkerAsync();
        }

        private void testButton_Click(object sender, RoutedEventArgs e)
        {
        }
    }
}
