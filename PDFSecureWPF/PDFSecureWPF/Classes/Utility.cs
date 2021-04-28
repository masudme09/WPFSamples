//using iTextSharp.text.pdf;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System;
using System.Collections.Generic;
using System.Drawing.Imaging;
using PDFLibNet32;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PDFSecureWPF.Classes
{
    public static class Utility
    {
        public static string Browse()
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog
            {
                InitialDirectory = @"D:\",
                Title = "Select PDF File",

                CheckFileExists = true,
                CheckPathExists = true,

                DefaultExt = ".pdf",
                Filter = "pdf files (*.pdf)|*.pdf",
                FilterIndex = 2,
                RestoreDirectory = true,

                ReadOnlyChecked = true,
                ShowReadOnly = true
            };

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                return openFileDialog1.FileName;
            }
            else
                return null;
        }

        public static string[] BrowsePdfFiles()
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog
            {
                Multiselect = true,
                InitialDirectory = @"D:\",
                Title = "Select PDF Files",

                CheckFileExists = true,
                CheckPathExists = true,

                DefaultExt = ".pdf",
                Filter = "pdf files (*.pdf)|*.pdf",
                FilterIndex = 2,
                RestoreDirectory = true,

                ReadOnlyChecked = true,
                ShowReadOnly = true
            };

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                return openFileDialog1.FileNames;
            }
            else
                return null;
        }

        public static string SaveFile()
        {
            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            saveFileDialog1.InitialDirectory = @"C:\";
            saveFileDialog1.Title = "Save Secured Pdf File";
            saveFileDialog1.CheckFileExists = false;
            saveFileDialog1.CheckPathExists = true;
            saveFileDialog1.DefaultExt = "pdf";
            saveFileDialog1.Filter = "Pdf files (*.pdf)|*.pdf";
            saveFileDialog1.FilterIndex = 2;
            saveFileDialog1.RestoreDirectory = true;
            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                return saveFileDialog1.FileName;
            }
            else
                return null;
        }


        public enum Definition
        {
            One = 1, Two = 2, Three = 3, Four = 4, Five = 5, Six = 6, Seven = 7, Eight = 8, Nine = 9, Ten = 10
        }

        /*
        public static void ConvertPDF2Image(string pdfInputPath, string imageOutputPath,
             int startPageNum, int endPageNum, ImageFormat imageFormat, Definition definition, int dpi)
        {

            using (var document = PdfDocument.(pdfInputPath))
            {

                if (!System.IO.Directory.Exists(imageOutputPath))
                {
                    System.IO.Directory.CreateDirectory(imageOutputPath);
                }

                // validate pageNum
                if (startPageNum <= 0)
                {
                    startPageNum = 1;
                }

                if (endPageNum > document.PageCount)
                {
                    endPageNum = document.PageCount;
                }

                if (startPageNum > endPageNum)
                {
                    int tempPageNum = startPageNum;
                    startPageNum = endPageNum;
                    endPageNum = startPageNum;
                }

                // start to convert each page
                for (int i = startPageNum; i <= endPageNum; i++)
                {
                    var bitmapImage = document.Render(0, 300, 300, true);
                    bitmapImage.Save(System.IO.Path.Combine(imageOutputPath, i.ToString() + ".jpg"), ImageFormat.Bmp);

                    //using (var image = document.Render(i, dpi, dpi, PdfRenderFlags.CorrectFromDpi))
                    //{
                    //    var encoder = ImageCodecInfo.GetImageEncoders().First(C => C.FormatID == ImageFormat.Jpeg.Guid);
                    //    var encParms = new EncoderParameters(1);
                    //    encParms.Param[0] = new EncoderParameter(System.Drawing.Imaging.Encoder.Quality, 100L);
                    //    image.Save(System.IO.Path.Combine(imageOutputPath, i.ToString() + ".jpg"), encoder, encParms);

                    //}

                }

            }
        }
        */

        
    }
}
