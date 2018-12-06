using Bytescout.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace BuildingTemplate
{
    public struct StringSet
    {
        public string title;
        public string url;        
    };

    class ManageExcel
    {
        public static int usedRows = 0;
        public static Spreadsheet document = null;
        public static Worksheet worksheet = null;

        public static List<StringSet> list = null;
        public static string outputfolder = null;

        public static bool SpreadSheetOepnAndLoad(string spreadsheeetFullPath)
        {
            // Create new Spreadsheet
            document = new Spreadsheet();
            document.LoadFromFile(spreadsheeetFullPath);

            // Get first worksheet
            worksheet = document.Workbook.Worksheets.ByName("Feuil1");
            list = new List<StringSet>();

            return true;
        }

        public static bool SpreadSheetSaveAndClose(string spreadsheeetFullPath)
        {
            // Save document
            document.SaveAs(spreadsheeetFullPath);

            //Close document
            document.Dispose();
            document.Close();

            return true;
        }

        public static int getusedrowcount()
        {
            string column_A = "Title";
            int nusedRow = 0;

            while (column_A != "")
            {
                column_A = worksheet.Cell(nusedRow, 0).ValueAsString;
                nusedRow++;

            }
            return nusedRow;
        }

        public static void processingImage(string text, string urlname, int ncurindex)
        {
            // Create a new image
            Image newimg = new Bitmap(600, 900);
            Image oldtemplate = Image.FromFile("SAMPLE.jpg");
            newimg = (Image)oldtemplate.Clone();

            Image insertedimage = Image.FromFile(urlname);

            Graphics g = Graphics.FromImage(newimg);
            g.SmoothingMode = SmoothingMode.AntiAlias;


            g.DrawImage(insertedimage, 0, 525, 600, 320);
            //g.DrawString("That's my boy!",   new Font("Arial", 12, FontStyle.Bold),   SystemBrushes.WindowText, new Point(5, 100));


            StringFormat sf = new StringFormat();
            sf.LineAlignment = StringAlignment.Center;
            sf.Alignment = StringAlignment.Center;

            //string test = "7 astuces naturelles pour vous débarrasser des rats de chez vous ... sans produits chimiques !";
            g.DrawString(text, new Font("Bahnschrift", 14, FontStyle.Bold), new SolidBrush(Color.White), new RectangleF(0, 100, 600, 400), sf);


            string savedname = outputfolder + "\\" + ncurindex.ToString() + ".jpg";


            newimg.Save(savedname, ImageFormat.Jpeg);
            g.Dispose();
            newimg.Dispose();

            //RemoveFileFromHarddesk(urlname);

        }

        public static void RemoveFileFromHarddesk(string filePath)
        {
            if (System.IO.File.Exists(filePath))

            {
                System.IO.File.Delete(filePath);

            }
        }

        public static void RemoveDirectory()
        {
            if (System.IO.Directory.Exists("urlimages"))

            {
                System.IO.Directory.Delete("urlimages", true);

            }
        }

        public static void CreateDirectory()
        {
            if (!System.IO.Directory.Exists("urlimages"))

            {
                System.IO.Directory.CreateDirectory("urlimages");

            }
        }

        public static void CreateOutputDirectory()
        {
            if (!System.IO.Directory.Exists(outputfolder))
            {
                System.IO.Directory.CreateDirectory(outputfolder);

            }
            
        }

        //System.IO.Directory.CreateDirectory(pathString);


        public static void SpreadSheetScanRow(string spreadsheeetFullPath)
        {

            outputfolder = "Output_" + DateTime.Now.ToString("HH_mm_ss");

            //MessageBox.Show(outputfolder);
            // Create a new image
            /*Image newimg = new Bitmap(600, 900);
            Image oldtemplate = Image.FromFile("D:\\temp\\sample.jpg");
            newimg = (Image)oldtemplate.Clone();*/

            CreateDirectory();
            CreateOutputDirectory();

            usedRows = getusedrowcount();

            string column_A = "";
            string column_B = "";

            int nIndex = 1;
            StringSet temp = new StringSet();

            string localFileName = "";

            while (usedRows-1 > nIndex)
            {
                column_A = worksheet.Cell(nIndex, 0).ValueAsString;
                column_B = worksheet.Cell(nIndex, 1).ValueAsString;

                temp.title = column_A;
                temp.url = column_B;

                //Console.WriteLine(column_A + " : " +  column_B);

                localFileName = "urlimages\\" + nIndex.ToString() + ".jpg";

                WebClient webClient = new WebClient();
                webClient.DownloadFile(column_B, localFileName);

                processingImage(column_A, localFileName, nIndex);

                int progressvalue = (int)(100*nIndex / usedRows);

                Console.WriteLine(progressvalue);

                Form1.m_progressBar.BeginInvoke(new Action(() => Form1.m_progressBar.Value = progressvalue));
                Form1.Label_Percent.Text = (progressvalue).ToString() + "%";

                Console.WriteLine(localFileName);
                //RemoveFileFromHarddesk(localFileName);

                list.Add(temp);

                nIndex++;

            }

            Thread.Sleep(5000);
            Form1.m_progressBar.BeginInvoke(new Action(() => Form1.m_progressBar.Value = 100));
            Form1.Label_Percent.Text = (100).ToString() + "%";

            MessageBox.Show("Successfully completed!");
            //RemoveDirectory();

        }
    }
}
