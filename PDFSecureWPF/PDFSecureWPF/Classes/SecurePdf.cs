using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using PDFSecureWPF;
using System.IO;
using System.ComponentModel;
using System.Drawing.Imaging;
using System.Drawing;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System.Threading;
using System.Windows.Forms;
using System.Diagnostics;
using System.Windows;
using PDFLibNet32;

namespace PDFSecureWPF.Classes
{
   
    public class SecurePdf
    {
        List<string> sourceFilesPaths;
        string saveFilePath;
        int dpi = MainWindow.selectedOutputQuality;
        string ZipPassword;
        bool convertToSingle;
        bool isBatch;
        string saveFilePathForConvertToSingle;

        /// <summary>
        /// This method collects values from mainwindow static variables
        /// </summary>
        public void getValuesFromMainWindow()
        {
            sourceFilesPaths=new List<string>();
            sourceFilesPaths = MainWindow.selectedFilePath;
            dpi = MainWindow.selectedOutputQuality;
        }

        /// <summary>
        /// For Single pdf secure
        /// </summary>
        /// <param name="outputQuality"></param>
        /// <param name="sourceFilesPaths"></param>
        public SecurePdf(int outputQuality, List<string> sourceFilesPaths, string ZipPassword)
        {
            this.dpi = outputQuality;
            this.sourceFilesPaths = sourceFilesPaths;
            this.ZipPassword = ZipPassword;
            this.convertToSingle = false;
            isBatch = false;
        }

        /// <summary>
        /// For batch secure
        /// </summary>
        /// <param name="outputQuality"></param>
        /// <param name="sourceFilesPaths"></param>
        /// <param name="ZipPassword"></param>
        /// <param name="convertToSingle"></param>
        public SecurePdf(int outputQuality, List<string> sourceFilesPaths, string ZipPassword,bool convertToSingle, string saveFilePath)
        {
            this.dpi = outputQuality;
            this.sourceFilesPaths = sourceFilesPaths;
            this.ZipPassword = ZipPassword;
            this.convertToSingle = convertToSingle;
            isBatch = true;
            saveFilePathForConvertToSingle = saveFilePath;
        }


        public void makeSecure(BackgroundWorker worker)
        {
            long oldCount=0; //to keep updated the old files count for makig single pdf

            #region creatingTempImageLocation
            //Creating Temp image location and first clear all if already there
            string tempLocation = Path.Combine(Path.GetTempPath(), "securepdf");

            if (Directory.Exists(tempLocation))
            {
                try
                {
                    Directory.Delete(tempLocation, true);
                }
                catch (Exception ex)
                {
                    System.Windows.MessageBox.Show(ex.ToString());
                }

            }

            Directory.CreateDirectory(tempLocation);

            string tempImageLocation = tempLocation;

            #endregion creatingTempImageLocation

            if (convertToSingle == true)
                saveFilePath = saveFilePathForConvertToSingle;

            foreach (string sourcePath in sourceFilesPaths)
            {
                if (!File.Exists(sourcePath))
                {
                    System.Windows.MessageBox.Show("Please select source pdf file first");
                    return;
                }

                if(isBatch==true)
                {
                    MainWindow.currentProcessingFile = "Processing: " + Path.GetFileName(sourcePath);
                }

                if(convertToSingle==false)
                    saveFilePath = Path.Combine(Path.GetDirectoryName(sourcePath), Path.GetFileName(sourcePath).Replace(".pdf", "Secured.pdf"));

                #region convertToSecureAndReportProgress
                PDFWrapper pdfWrapper = new PDFWrapper();

                pdfWrapper.LoadPDF(sourcePath);

                //Reporting progress
                worker.ReportProgress(1);

                long pageCount = (long) pdfWrapper.PageCount;
                
                int startIndex=0;
                long endIndex=pageCount;

                if(convertToSingle==true) //for converting to single pdf
                {
                    startIndex = (int) oldCount;
                    oldCount = oldCount + pageCount;
                    pageCount = oldCount;
                }
                
                float progressUnit = (float)100.0 / (pageCount * 2);
                float reportProgress = 1;

                #region convertingToImages
                string directoryOfFile = Path.GetDirectoryName(saveFilePath);                            

                for (var i = startIndex; i <= endIndex; i++)
                {
                    try
                    {
                        //Utility.ConvertPDF2Image(sourcePath, tempImageLocation + "\\", i, i, ImageFormat.Jpeg, Utility.Definition.One, dpi);
                        reportProgress = reportProgress + progressUnit;

                        //Reporting progress
                        worker.ReportProgress((int)reportProgress);

                    }
                    catch (Exception ex)
                    {
                        System.Windows.MessageBox.Show(ex.ToString());

                    }
                }
                #endregion convertingToImages

                if(convertToSingle==false) //if convertToSingle is false then make individual pdf else make all image and then convert all of them to single pdf
                {

                    convertingImagesToPdf(pageCount, tempImageLocation, reportProgress, progressUnit, worker);
                }               


                //Finishing operation and file open dialogue
                worker.ReportProgress(100);
                Thread.Sleep(50);

                //if only one file in the list then show dialogue to open that
                if(isBatch==false)
                {
                    System.Windows.MessageBoxResult messageBoxResult = System.Windows.MessageBox.Show("Your pdf is now secured, do you want to open it?", "Secured!",System.Windows.MessageBoxButton.YesNo,
                        System.Windows.MessageBoxImage.None,MessageBoxResult.None,System.Windows.MessageBoxOptions.DefaultDesktopOnly);
                    if (messageBoxResult == MessageBoxResult.Yes)
                        Process.Start(saveFilePath);
                }

                #endregion convertToSecureAndReportProgress
            }

           
            if (isBatch==true && convertToSingle==true)
            {
                float reportProgress = 1;
                float progressUnit = (float)100.0 / (oldCount * 2);
                convertingImagesToPdf(oldCount, tempImageLocation,reportProgress, progressUnit, worker);
            }
            else if (isBatch == true)
            {

                System.Windows.MessageBox.Show("Your pdf files are now secured.", "Secured!", System.Windows.MessageBoxButton.OK,
                    System.Windows.MessageBoxImage.None, MessageBoxResult.None, System.Windows.MessageBoxOptions.DefaultDesktopOnly);
            }

        }


        public void convertingImagesToPdf(long imageCount, string tempImageLocation, float reportProgress, float progressUnit, BackgroundWorker worker)
        {
            //create individual pdf
            #region convertingImagesBackToPDF
            using (var stream = new FileStream(saveFilePath, FileMode.Create, FileAccess.Write, FileShare.None))
            {

                iTextSharp.text.Rectangle pageSize = null;
                if (imageCount > 0)
                {
                    using (var srcImage = new Bitmap(tempImageLocation + @"\" + 1 + ".jpg"))
                    {
                        pageSize = new iTextSharp.text.Rectangle(0, 0, srcImage.Width, srcImage.Height);
                    }
                }
                else
                {
                    return;
                }
                Document pdfDocToWrite = new Document(pageSize, 0, 0, 10, 10);
                PdfWriter.GetInstance(pdfDocToWrite, stream);
                pdfDocToWrite.Open();
                for (var i = 1; i <= imageCount; i++)
                {

                    using (var srcImage = new Bitmap(tempImageLocation + @"\" + i + ".jpg"))
                    {
                        pageSize = new iTextSharp.text.Rectangle(0, 0, srcImage.Width, srcImage.Height);
                    }

                    //if(i==0)
                    pdfDocToWrite.SetPageSize(pageSize);
                    pdfDocToWrite.SetMargins(0, 0, 10, 10);
                    pdfDocToWrite.SetMargins(0, 0, 10, 10);

                    if (i > 0)
                        pdfDocToWrite.NewPage();

                    using (var imageStream = new FileStream(tempImageLocation + @"\" + i + ".jpg", FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                    {
                        var image1 = iTextSharp.text.Image.GetInstance(imageStream);
                        pdfDocToWrite.Add(image1);

                        reportProgress = reportProgress + progressUnit;
                        //Reporting progress
                        worker.ReportProgress((int)reportProgress);

                    }

                }
                pdfDocToWrite.Close();
            }
            #endregion convertingImagesBackToPDF

            //Delete all images on Temp location
            if (imageCount > 0)
            {
                for (var i = 1; i <= imageCount; i++)
                {
                    File.Delete(tempImageLocation + @"\" + i + ".jpg");
                }
            }
        }

        public void reportProgress(System.Windows.Controls.ProgressBar prBar, ProgressChangedEventArgs e)
        {
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

    }
}
