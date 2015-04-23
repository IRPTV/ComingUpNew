using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Text;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using ComingUp.MyDBTableAdapters;
using System.Configuration;
using System.Diagnostics;
using System.IO;

namespace ComingUp
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        private void button1_Click(object sender, EventArgs e)
        {

            string[] FilesList = Directory.GetFiles(ConfigurationSettings.AppSettings["OutputPath"].ToString().Trim());
            foreach (string item in FilesList)
            {
                try
                {
                    if (File.GetLastAccessTime(item) < DateTime.Now.AddHours(-48))
                    {
                        File.Delete(item);
                        richTextBox1.Text += (item) + " *Deleted* \n";
                        richTextBox1.SelectionStart = richTextBox1.Text.Length;
                        richTextBox1.ScrollToCaret();
                        Application.DoEvents();
                    }
                }
                catch (Exception Exp)
                {
                    richTextBox1.Text += (Exp) + " \n";
                    richTextBox1.SelectionStart = richTextBox1.Text.Length;
                    richTextBox1.ScrollToCaret();
                    Application.DoEvents();
                }
               
            }


            button1.ForeColor = Color.White;
            button1.Text = "Started";
            button1.BackColor = Color.Red;

            richTextBox1.Text = "";

            SelectSchedules();
            Renderer();


          
        }
        public static Bitmap Convert_Text_to_Image(string txt, string fontname, int fontsize)
        {
            //creating bitmap image
            Bitmap bmp = new Bitmap(1, 1);

            //FromImage method creates a new Graphics from the specified Image.
            Graphics graphics = Graphics.FromImage(bmp);
            // Create the Font object for the image text drawing.
            Font font = new Font(fontname, fontsize, FontStyle.Bold);
            // Instantiating object of Bitmap image again with the correct size for the text and font.
            SizeF stringSize = graphics.MeasureString(txt, font);
            // bmp = new Bitmap(bmp, (int)stringSize.Width, (int)stringSize.Height);
            bmp = new Bitmap(bmp, 1160, 110);
            graphics = Graphics.FromImage(bmp);

            /* It can also be a way
           bmp = new Bitmap(bmp, new Size((int)graphics.MeasureString(txt, font).Width, (int)graphics.MeasureString(txt, font).Height));*/

            //Draw Specified text with specified format 

            graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality;
            graphics.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit; // <-- important!
            graphics.CompositingQuality = System.Drawing.Drawing2D.CompositingQuality.HighQuality;
            graphics.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
            graphics.PixelOffsetMode = System.Drawing.Drawing2D.PixelOffsetMode.HighQuality;
            graphics.TextContrast = 0;
            graphics.DrawString(txt, font, Brushes.White, 50, 50);
            font.Dispose();
            graphics.Flush();
            graphics.Dispose();
            return bmp;     //return Bitmap Image 
        }
        private Bitmap CreateBitmapImage(string sImageText)
        {
            Bitmap objBmpImage = new Bitmap(1, 1);

            int intWidth = 0;
            int intHeight = 0;

            // Create the Font object for the image text drawing.
            Font objFont = new Font("Context Reprise BlackExp SSi", 30, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Pixel);

            // Create a graphics object to measure the text's width and height.
            Graphics objGraphics = Graphics.FromImage(objBmpImage);

            // This is where the bitmap size is determined.
            intWidth = (int)objGraphics.MeasureString(sImageText, objFont).Width;
            intHeight = (int)objGraphics.MeasureString(sImageText, objFont).Height;

            // Create the bmpImage again with the correct size for the text and font.
            objBmpImage = new Bitmap(objBmpImage, new Size(1160, 110));

            // Add the colors to the new bitmap.
            objGraphics = Graphics.FromImage(objBmpImage);

            // Set Background color
            //objGraphics.Clear(Color.Black);
            objGraphics.SmoothingMode = SmoothingMode.AntiAlias;
            objGraphics.TextRenderingHint = TextRenderingHint.AntiAlias;
            objGraphics.DrawString(sImageText, objFont, Brushes.White, 50, 50);
            objGraphics.Flush();

            return (objBmpImage);
        }
        protected void GenerateTitleImage(string Title, int FileIndex)
        {
            float fontSize = int.Parse(ConfigurationSettings.AppSettings["TitleFontSize"].ToString().Trim());

            //some test image for this demo
            Bitmap bmp = new Bitmap(1160, 110);
            Graphics g = Graphics.FromImage(bmp);

            //this will center align our text at the bottom of the image
            StringFormat sf = new StringFormat();
            //sf.Alignment = StringAlignment.Near;
            //sf.LineAlignment = StringAlignment.Far;

            //define a font to use.
            Font f = new Font("Context Reprise SSi", fontSize, FontStyle.Bold, GraphicsUnit.Pixel);

            //pen for outline - set width parameter
            Pen p = new Pen(ColorTranslator.FromHtml("#000000"), 3);
            p.LineJoin = LineJoin.Round; //prevent "spikes" at the path

            //this makes the gradient repeat for each text line
            //Rectangle fr = new Rectangle(0, bmp.Height - f.Height, bmp.Width, f.Height);
            Rectangle fr = new Rectangle(0, 0, bmp.Width, f.Height);
            LinearGradientBrush b = new LinearGradientBrush(fr,
                                                            ColorTranslator.FromHtml("#ffffff"),
                                                            ColorTranslator.FromHtml("#ffffff"),
                                                            90);

            //this will be the rectangle used to draw and auto-wrap the text.
            //basically = image size
            Rectangle r = new Rectangle(2, 25, bmp.Width, bmp.Height);

            GraphicsPath gp = new GraphicsPath();

            //look mom! no pre-wrapping!
            gp.AddString(Title,
                         f.FontFamily, (int)f.Style, fontSize, r, sf);

            //these affect lines such as those in paths. Textrenderhint doesn't affect
            //text in a path as it is converted to ..well, a path.    
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.PixelOffsetMode = PixelOffsetMode.HighQuality;


            g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality;
            g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit; // <-- important!
            g.CompositingQuality = System.Drawing.Drawing2D.CompositingQuality.HighQuality;
            g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
            g.PixelOffsetMode = System.Drawing.Drawing2D.PixelOffsetMode.HighQuality;
            g.TextContrast = 0;
            //TODO: shadow -> g.translate, fillpath once, remove translate


            //var matrix = new Matrix();
            //matrix.Translate(10, 10);
            //g.Transform=matrix;
            g.DrawPath(p, gp);
            g.FillPath(b, gp);

            //cleanup
            gp.Dispose();
            b.Dispose();
            b.Dispose();
            f.Dispose();
            sf.Dispose();
            g.Dispose();

            bmp.Save(ConfigurationSettings.AppSettings["ImagePathTitle"].ToString().Trim() + string.Format("{0:00}", FileIndex) + ".png", System.Drawing.Imaging.ImageFormat.Png);
            bmp.Dispose();
        }
        protected void GenerateTimeImage(string TimeText, int FileIndex)
        {
            float fontSize = 46;

            //some test image for this demo
            Bitmap bmp = new Bitmap(1160, 110);
            Graphics g = Graphics.FromImage(bmp);

            //this will center align our text at the bottom of the image
            StringFormat sf = new StringFormat();
            //sf.Alignment = StringAlignment.Near;
            //sf.LineAlignment = StringAlignment.Far;

            //define a font to use.
            Font f = new Font("Context Reprise SSi", fontSize, FontStyle.Bold, GraphicsUnit.Pixel);

            //pen for outline - set width parameter
            Pen p = new Pen(ColorTranslator.FromHtml("#000000"), 3);
            p.LineJoin = LineJoin.Round; //prevent "spikes" at the path

            //this makes the gradient repeat for each text line
            //Rectangle fr = new Rectangle(0, bmp.Height - f.Height, bmp.Width, f.Height);
            Rectangle fr = new Rectangle(0, 0, bmp.Width, f.Height);
            LinearGradientBrush b = new LinearGradientBrush(fr,
                                                            ColorTranslator.FromHtml("#ffffff"),
                                                            ColorTranslator.FromHtml("#ffffff"),
                                                            90);

            //this will be the rectangle used to draw and auto-wrap the text.
            //basically = image size
            Rectangle r = new Rectangle(2, 25, bmp.Width, bmp.Height);

            GraphicsPath gp = new GraphicsPath();

            //look mom! no pre-wrapping!
            gp.AddString(TimeText,
                         f.FontFamily, (int)f.Style, fontSize, r, sf);

            //these affect lines such as those in paths. Textrenderhint doesn't affect
            //text in a path as it is converted to ..well, a path.    
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.PixelOffsetMode = PixelOffsetMode.HighQuality;


            g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality;
            g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit; // <-- important!
            g.CompositingQuality = System.Drawing.Drawing2D.CompositingQuality.HighQuality;
            g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
            g.PixelOffsetMode = System.Drawing.Drawing2D.PixelOffsetMode.HighQuality;
            g.TextContrast = 0;

            g.DrawPath(p, gp);
            g.FillPath(b, gp);

            //cleanup
            gp.Dispose();
            b.Dispose();
            b.Dispose();
            f.Dispose();
            sf.Dispose();
            g.Dispose();

            bmp.Save(ConfigurationSettings.AppSettings["ImagePathTime"].ToString().Trim() + "Time " + string.Format("{0:00}", FileIndex) + ".png", System.Drawing.Imaging.ImageFormat.Png);
            bmp.Dispose();
        }
        protected void SelectSchedules()
        {
            timer1.Enabled = false;
            SchedulesDataTableAdapter Sch_Ta = new SchedulesDataTableAdapter();
            MyDB.SchedulesDataTable Sch_Dt = Sch_Ta.Select_Top(DateTime.Now);
            for (int i = 0; i < Sch_Dt.Rows.Count; i++)
            {
                string ProgName = Sch_Dt.Rows[i]["name"].ToString();

                //2014-01-25 Replace Documentry by Doc
                ProgName = ProgName.Replace("documentary", "Doc");
                ProgName = ProgName.Replace("Documentary", "Doc");


                int FirstIndex = ProgName.IndexOf("-");
                int SecondIndex = 0;
                if (FirstIndex > 0)
                {
                    SecondIndex = ProgName.IndexOf("-", FirstIndex + 1);
                    if (SecondIndex > FirstIndex)
                    {
                        ProgName = ProgName.Remove(FirstIndex, SecondIndex - FirstIndex + 1);
                        ProgName = ProgName.Insert(FirstIndex, ":");
                    }
                    ProgName = ProgName.Replace("  :", ":");
                    ProgName = ProgName.Replace(" :", ":");

                }
                int TitleLenght = int.Parse(ConfigurationSettings.AppSettings["TitleLenght"].ToString().Trim());
                if (ProgName.Length > TitleLenght)
                {
                    char[] delimiters = new char[] { ' ' };
                    string[] PrgNameList = ProgName.Split(delimiters, StringSplitOptions.RemoveEmptyEntries);
                    int CutIndex = 0;
                    string OutName = "";
                    bool AllowAdd = true;
                    foreach (string item in PrgNameList)
                    {
                        if (CutIndex + item.Length + 1 < TitleLenght)
                        {
                            if (AllowAdd)
                            {
                                CutIndex += item.Length + 1;
                                OutName += item + " ";
                            }

                        }
                        else
                        {
                            AllowAdd = false;
                        }
                    }
                    //ProgName = ProgName.Remove(CutIndex, ProgName.Length - CutIndex);
                    //ProgName += "...";

                    ProgName = OutName + "...";
                }


                GenerateTitleImage(ProgName, i + 1);
                DateTime ProgTime = DateTime.Parse(Sch_Dt.Rows[i]["datetime"].ToString());
                string ProgTimeText = ProgTime.ToShortTimeString().Substring(0, 5);

                GenerateTimeImage(string.Format("{0:00}", ProgTime.Hour) + ":" + ConfigurationSettings.AppSettings["TimeScheduleMinute"].ToString().Trim(), i + 1);
                richTextBox1.Text += ProgName + "  (  " + string.Format("{0:00}", ProgTime.Hour) + ":" + ConfigurationSettings.AppSettings["TimeScheduleMinute"].ToString().Trim() + "  )  \n";
                richTextBox1.SelectionStart = richTextBox1.Text.Length;
                richTextBox1.ScrollToCaret();
                Application.DoEvents();
            }



        }
        private void timer1_Tick(object sender, EventArgs e)
        {
            int StartMinute = int.Parse(ConfigurationSettings.AppSettings["RenderIntervalMin"].ToString().Trim());
            if (DateTime.Now.Minute >= StartMinute && DateTime.Now.Minute <= StartMinute + 1)
            {
                timer1.Enabled = false;
                button1_Click(new object(), new EventArgs());                
            }



        }
        private void Form1_Load(object sender, EventArgs e)
        {

            //timer1.Interval = int.Parse(ConfigurationSettings.AppSettings["RenderIntervalMin"].ToString().Trim()) * 60 * 1000;
        }
        protected void Renderer()
        {
            Process proc = new Process();
            proc.StartInfo.FileName = "\"" + ConfigurationSettings.AppSettings["AeRenderPath"].ToString().Trim() + "\"";
            string DateTimeStr = string.Format("{0:0000}", DateTime.Now.Year) + "-" + string.Format("{0:00}", DateTime.Now.Month) + "-" + string.Format("{0:00}", DateTime.Now.Day) + "_" + string.Format("{0:00}", DateTime.Now.Hour) + "-" + string.Format("{0:00}", DateTime.Now.Minute) + "-" + string.Format("{0:00}", DateTime.Now.Second);

            DirectoryInfo Dir = new DirectoryInfo(ConfigurationSettings.AppSettings["OutputPath"].ToString().Trim());

            if (!Dir.Exists)
            {
                Dir.Create();
            }


            proc.StartInfo.Arguments = " -project " + "\"" + ConfigurationSettings.AppSettings["AeProjectFile"].ToString().Trim() + "\"" + "   -comp   \"" + ConfigurationSettings.AppSettings["Composition"].ToString().Trim() + "\" -output " + "\"" + ConfigurationSettings.AppSettings["OutputPath"].ToString().Trim() + ConfigurationSettings.AppSettings["OutPutFileName"].ToString().Trim() + "_" + DateTimeStr + ".mp4" + "\"";
            proc.StartInfo.RedirectStandardError = true;
            proc.StartInfo.UseShellExecute = false;
            proc.StartInfo.CreateNoWindow = true;
            proc.EnableRaisingEvents = true;
            proc.StartInfo.RedirectStandardOutput = true;
            proc.StartInfo.RedirectStandardError = true;

            if (!proc.Start())
            {
                return;
            }

            proc.PriorityClass = ProcessPriorityClass.Normal;
            StreamReader reader = proc.StandardOutput;
            string line;
            while ((line = reader.ReadLine()) != null)
            {
                if (richTextBox1.Lines.Length > 8)
                {
                    richTextBox1.Text = "";
                }
                richTextBox1.Text += (line) + " \n";
                richTextBox1.SelectionStart = richTextBox1.Text.Length;
                richTextBox1.ScrollToCaret();
                Application.DoEvents();
            }
            proc.Close();

            try
            {
                string StaticDestFileName = ConfigurationSettings.AppSettings["ScheduleDestFileName"].ToString().Trim();
               // File.Delete(StaticDestFileName);
                File.Copy(ConfigurationSettings.AppSettings["OutputPath"].ToString().Trim() + ConfigurationSettings.AppSettings["OutPutFileName"].ToString().Trim() + "_" + DateTimeStr + ".mp4", StaticDestFileName,true);
                richTextBox1.Text += "COPY FINAL:"+StaticDestFileName +" \n";
                richTextBox1.SelectionStart = richTextBox1.Text.Length;
                richTextBox1.ScrollToCaret();
                Application.DoEvents();
            }
            catch (Exception Ex)
            {
                richTextBox1.Text += Ex.Message + " \n";
                richTextBox1.SelectionStart = richTextBox1.Text.Length;
                richTextBox1.ScrollToCaret();
                Application.DoEvents();
            }


            timer1.Enabled = true;
            this.Text = "CmgUp V1.7 2014-01-25: " + DateTime.Now.ToString();
            button1.ForeColor = Color.White;
            button1.Text = "Start";
            button1.BackColor = Color.Navy;
        }

    }
}
