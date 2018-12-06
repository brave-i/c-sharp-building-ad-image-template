using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using System.Drawing.Imaging;
using System.Drawing.Drawing2D;
using System.Threading;
using System.Net;

namespace BuildingTemplate
{
    public partial class Form1 : Form
    {

        public static string Excel_Path_String = null;
        private bool isProcessRunning = false;

        public Form1()
        {
            InitializeComponent();
        }

        public void InsertingImage()
        {

            int DEFAULT_WIDTH = 600;
            int DEFAULT_HEIGHT = 320;
            // Create a new image
            Image newimg = new Bitmap(DEFAULT_WIDTH, DEFAULT_HEIGHT);

            Image oldtemplate = Image.FromFile("D:\\temp\\sample.jpg");
            newimg = (Image)oldtemplate.Clone();

            Image insertedimage = Image.FromFile("D:\\temp\\2.jpg");
            //Image insertedimage = Image.FromFile("https://astucescuisine.com/composants/uploads/2017/03/089f1643f80b7d950ae91ee5a70fa189.jpg");
            newimg = (Image)oldtemplate.Clone();


            Graphics g = Graphics.FromImage(newimg);
            g.SmoothingMode = SmoothingMode.AntiAlias;


            g.DrawImage(insertedimage, 0, 525, 600, 320);
            //g.DrawString("That's my boy!",   new Font("Arial", 12, FontStyle.Bold),   SystemBrushes.WindowText, new Point(5, 100));


            StringFormat sf = new StringFormat();
            sf.LineAlignment = StringAlignment.Center;
            sf.Alignment = StringAlignment.Center;

            string test = "7 astuces naturelles pour vous débarrasser des rats de chez vous ... sans produits chimiques !";
            g.DrawString(test, new Font("Bahnschrift", 14, FontStyle.Bold), new SolidBrush(Color.White), new RectangleF(0, 100, 600, 400), sf);

            newimg.Save("output.jpg", ImageFormat.Jpeg);
            g.Dispose();


        }
        public StringBuilder getStringList(string text)
        {
            StringBuilder builder = new StringBuilder();
            int MAX_ROW_LEN = 25;

            string[] firstwords = text.Split(':');
            string firstrow_text = "";

            if (firstwords.Length == 1)//have no ":" character...
            {
                string[] words = firstwords[0].Split(' ');

                int count = 0;
                string subtext = "";
                string prevsubtext = "";

                while (words.Length - 1 > count)
                {

                    subtext = subtext + words[count] + " ";
                    if (subtext.Length > MAX_ROW_LEN)
                    {
                        subtext = "";
                        builder.Append(prevsubtext).AppendLine();
                    }

                    else
                    {
                        prevsubtext = subtext;
                        count++;
                    }

                    if (words.Length - 1 == count)
                    {
                        builder.Append(subtext).AppendLine();
                    }
                }
            }
            else
            {
                firstrow_text = firstwords[0];
                string[] words1 = firstrow_text.Split(' ');
                int count1 = 0;
                string subtext1 = "";
                string prevsubtext1 = "";

                if (firstrow_text.Length > MAX_ROW_LEN)
                {
                    while (words1.Length - 1 > count1)
                    {

                        subtext1 = subtext1 + words1[count1] + " ";
                        if (subtext1.Length > MAX_ROW_LEN)
                        {
                            subtext1 = "";
                            builder.Append(prevsubtext1).AppendLine();
                        }

                        else
                        {
                            prevsubtext1 = subtext1;
                            count1++;
                        }

                        if (words1.Length - 1 == count1)
                        {
                            builder.Append(subtext1 + ":").AppendLine();
                        }

                    }
                }
                else
                {
                    builder.Append(firstrow_text + ":").AppendLine();
                }
                //MessageBox.Show(builder.ToString().Split('\n').Length.ToString());


                string[] words = firstwords[1].Split(' ');

                int count = 0;
                string subtext = "";
                string prevsubtext = "";

                while (words.Length - 1 > count)
                {

                    subtext = subtext + words[count] + " ";
                    if (subtext.Length > MAX_ROW_LEN)
                    {
                        subtext = "";
                        builder.Append(prevsubtext).AppendLine();
                    }

                    else
                    {
                        prevsubtext = subtext;
                        count++;
                    }

                    if (words.Length - 1 == count)
                    {
                        builder.Append(subtext).AppendLine();
                    }
                }
            }
            //MessageBox.Show(builder.ToString());


            return builder;
        }

        private void Excel_Path_Btn_Click(object sender, EventArgs e)
        {
            //string path = "";
            string defaultText = "Select Spread Sheet File!";

            OpenFileDialog ofd = new OpenFileDialog();

            ofd.CheckFileExists = true;
            ofd.CheckPathExists = true;

            ofd.DefaultExt = "xlsx";
            ofd.Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
            ofd.FilterIndex = 1;

            Excel_Path_Edit.Clear();

            if (ofd.ShowDialog() == DialogResult.OK)
            {
                Excel_Path_String = ofd.FileName;
                //Console.WriteLine(excelurl);                        //excel path

                Excel_Path_Edit.SelectedText = Excel_Path_String;


            }

            if (Excel_Path_String == null) Excel_Path_Edit.SelectedText = defaultText;
            else
            {
                Excel_Path_Edit.Clear();
                Excel_Path_Edit.SelectedText = Excel_Path_String;
            }
            //MessageBox.Show(Excel_Path_String);
            Console.WriteLine(Excel_Path_String);
        }

        public void RemoveDirectory()
        {
            if (System.IO.Directory.Exists("urlimages"))

            {
                System.IO.Directory.Delete("urlimages", true);

            }
        }

        private void m_startButton_Click(object sender, EventArgs e)
        {
            //MessageBox.Show(DateTime.Now.ToString("HH:mm:ss"));

            //return;

            //InsertingImage();
            //string remoteFileUrl = "https://astucescuisine.com/composants/uploads/2017/03/27811a9563490451c506d1e556fcf195.jpg";
            //string localFileName = "urlimages\\test.jpg";

            //WebClient webClient = new WebClient();
            //webClient.DownloadFile(remoteFileUrl, localFileName);

            //System.IO.Directory.Delete("abc", true);



            if (Excel_Path_String == null)
            {
                MessageBox.Show("Please Select Excel File!!");
                return;
            }

            if (isProcessRunning)
            {
                MessageBox.Show("A process is already running.");
                return;
            }

            /*ManageExcel.SpreadSheetOepnAndLoad(Excel_Path_String);
            ManageExcel.SpreadSheetScanRow(Excel_Path_String);
            ManageExcel.SpreadSheetSaveAndClose(Excel_Path_String);*/

            
            Thread backgroundThread = new Thread(
                new ThreadStart(() =>
                {
                    isProcessRunning = true;

                    ManageExcel.SpreadSheetOepnAndLoad(Excel_Path_String);
                    ManageExcel.SpreadSheetScanRow(Excel_Path_String);

                    //m_progressBar.BeginInvoke(new Action(() => Form1.m_progressva.Value = progressvalue));

                    ManageExcel.SpreadSheetSaveAndClose(Excel_Path_String);

                    isProcessRunning = false;
                    RemoveDirectory();

                }
            ));

            // Start the background process thread
            backgroundThread.Start();
            
        }
    }
}
