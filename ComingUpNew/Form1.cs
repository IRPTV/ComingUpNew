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

            try
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
            }
            catch { }

            button1.ForeColor = Color.White;
            button1.Text = "Started";
            button1.BackColor = Color.Red;

            richTextBox1.Text = "";
            SelectSchedules();
            Renderer();          
        }     
        protected void SelectSchedules()
        {
            timer1.Enabled = false;
            SchedulesDataTableAdapter Sch_Ta = new SchedulesDataTableAdapter();
            MyDB.SchedulesDataTable Sch_Dt = Sch_Ta.Select_Top(DateTime.Now);
            StringBuilder Data = new StringBuilder();
            for (int i = 0; i < Sch_Dt.Rows.Count; i++)
            {
                string ProgName = Sch_Dt.Rows[i]["name"].ToString();

                //2014-01-25 Replace Documentry by Doc
                ProgName = ProgName.Replace("documentary", "Doc");
                ProgName = ProgName.Replace("Documentary", "Doc");
                ProgName = ProgName.Replace("Medium Items - 1 - ", "");
                ProgName = ProgName.Replace("Medium Items - 2 - ", "");


                int FirstIndex = ProgName.IndexOf("-");
                int SecondIndex = 0;
                if (FirstIndex > 0)
                {
                    if (ProgName.StartsWith("Doc"))
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
                    else
                    {
                        //2015-07-13
                        ProgName = ProgName.Remove(FirstIndex, ProgName.Length-FirstIndex);
                    }

                }

                //Commented for new Coming UP:
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
                DateTime ProgTime = DateTime.Parse(Sch_Dt.Rows[i]["datetime"].ToString());
                //Program1 = ["The World after Fukushima 2","01:00"]

                //string ProgTimeText = ProgTime.Hour.ToString("00") + ":" + ConfigurationSettings.AppSettings["TimeScheduleMinute"].ToString();
                string ProgTimeText = ProgTime.Hour.ToString("00") + ":" + Math.Floor((decimal)ProgTime.Minute+1).ToString("00");
                Data.AppendLine("Program" + (i + 1).ToString() + " = [\"" + ProgName + "\",\"" + ProgTimeText + "\"]");

            }
            StreamWriter StrW = new StreamWriter(ConfigurationSettings.AppSettings["XmlDataFile"].ToString().Trim());
            StrW.Write(Data);
            StrW.Close();
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
            //this.Text = "CmgUpNew V1.7 2014-01-25: " + DateTime.Now.ToString();
            button1.ForeColor = Color.White;
            button1.Text = "Start";
            button1.BackColor = Color.Navy;
        }
    }
}
