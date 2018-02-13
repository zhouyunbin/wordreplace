using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using DocumentFormat.OpenXml.Wordprocessing;
using System.IO;
using System.Threading;
using System.Reflection;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;

namespace WordReplace
{
    public partial class Form1 : Form
    {
        SynchronizationContext m_SyncContext = null;
        ConfigFile config;
        public Form1()
        {
            InitializeComponent();
            config = ConfigFile.LoadOrCreateFile("config.ini");
            m_SyncContext = SynchronizationContext.Current;
            textBox1.Text = config.GetConfigValue("txt1");
            textBox2.Text = config.GetConfigValue("txt2");
            textBox3.Text = config.GetConfigValue("txt3");
            textBox4.Text = config.GetConfigValue("txt4");
            textBox5.Text = config.GetConfigValue("txt5");
            textBox6.Text = config.GetConfigValue("txt6");
            textBox7.Text = config.GetConfigValue("txt7");
            textBox8.Text = config.GetConfigValue("txt8");

        }

        private void SetTextSafePost(object text)
        {
            listBox1.Items.Add(text);
            this.listBox1.SelectedIndex = this.listBox1.Items.Count - 1;
            this.listBox1.SelectedIndex = -1;
        }

        private void ProgressPlus(object text)
        {
            progressBar1.Value++;
            if (progressBar1.Value == progressBar1.Maximum)
                MessageBox.Show("替换完成");
        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog path = new FolderBrowserDialog();
            path.ShowDialog();
            this.textBox3.Text = path.SelectedPath;
        }
        int threadcount;
        private readonly object Locker = new object();
        private void button1_Click(object sender, EventArgs e)
        {
            DirectoryInfo theFolder = new DirectoryInfo(textBox3.Text);
            progressBar1.Maximum = theFolder.GetFiles("*.doc*").Length;
            progressBar1.Value = 0;
            Thread mythread = new Thread(new ThreadStart(StartJob));
            mythread.Start();
        }

        private void StartJob()
        {
            threadcount = 0;
            DirectoryInfo theFolder = new DirectoryInfo(textBox3.Text);
            foreach (FileInfo NextFile in theFolder.GetFiles())
            {
                if (NextFile.Extension.Equals(".doc") || NextFile.Extension.Equals(".docx"))
                {
                    if(NextFile.Name.Contains("$"))
                    {
                        m_SyncContext.Post(ProgressPlus, null);
                        continue;
                    }
                    m_SyncContext.Post(SetTextSafePost, NextFile.Name);
                    
                    lock (Locker)
                    {
                        threadcount++;
                    }
                    while (threadcount > 10)
                    {
                        Thread.Sleep(100);
                    }
                    Thread mythread = new Thread(new ParameterizedThreadStart(Replace));
                    Console.WriteLine(NextFile.FullName);
                    mythread.Start(NextFile.FullName);

                }
            }
        }
        
        
        
        private void Replace(object filename)
        {
            //创建word

            
            try
            {
                //创建word应用程序
                if(filename.ToString().EndsWith("doc"))
                {
                    //SaveDocx(filename.ToString());
                   // filename = filename.ToString() + "x";
                   // m_SyncContext.Post(ProgressPlus, null);
                }
                WordprocessingDocument doc = WordprocessingDocument.Open(filename.ToString(), true);
                Body body = doc.MainDocumentPart.Document.Body;
                foreach (Paragraph paragraph in body.Elements<Paragraph>())
                {
                    if (checkBox6.Checked)
                    {
                        ReplaceAB(paragraph);
                    }
                    if (checkBox2.Checked)
                    {
                        DeleteBefore(paragraph);
                    }
                    if (checkBox7.Checked)
                    {
                        DeleteAfter(paragraph);
                    }
                    
                    foreach (Run run in paragraph.Elements<Run>())
                    {
                        if (checkBox1.Checked)
                        {
                            ReplaceAandB(run);
                        }
                        
                        if (checkBox3.Checked)
                        {
                            DeleteAlpha(run);
                        }
                        if (checkBox4.Checked)
                        {
                            DeleteNumber(run);
                        }
                        if (checkBox5.Checked)
                        {
                            DeleteLink(run);
                        }


                    }
                }
                    
                
                doc.Save();
                doc.Dispose();

                
                m_SyncContext.Post(ProgressPlus, null);
            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.Message);
                m_SyncContext.Post(ProgressPlus, null);
            }
            finally
            {
                lock (Locker)
                {
                    threadcount--;
                }
                
                
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            System.Environment.Exit(0);
        }

        private void button4_Click(object sender, EventArgs e)
        {
            

        }

        private void progressBar1_Click(object sender, EventArgs e)
        {

        }

        private void button4_Click_1(object sender, EventArgs e)
        {
            listBox1.Items.Clear();
        }
        protected override void OnClosed(EventArgs e)
        {
            base.OnClosed(e);
            config["txt1"] = textBox1.Text;
            config["txt2"] = textBox2.Text;
            config["txt3"] = textBox3.Text;
            config["txt4"] = textBox4.Text;
            config["txt5"] = textBox5.Text;
            config["txt6"] = textBox6.Text;
            config["txt7"] = textBox7.Text;
            config["txt8"] = textBox8.Text;
        }

        void ReplaceAandB(Run run)
        {
            string a = textBox4.Text;
            string b = textBox5.Text;
            if(run.InnerText.Contains(a))
            {
                string temp = run.InnerText.Replace(a,b);
                Text text=run.Elements<Text>().First();
                text.Remove();
                run.Append(new Text(temp));
            }
        }

        void ReplaceAB(Paragraph para)
        {
            string a = textBox8.Text;
            string b = textBox7.Text;
            ;
            if(para.InnerText.Contains(a)&&para.InnerText.Contains(b))
            {
                int flag = 0;
                List<Run> list = new List<Run>();
                foreach(Run r in para.Elements<Run>())
                {
                    if(r.InnerText.Contains(a)&&flag==0)
                    {
                        if(r.InnerText.Contains(b))
                        {
                            string s1= r.InnerText.Substring(0, r.InnerText.IndexOf(a) + a.Length);
                            string s2 = r.InnerText.Substring(r.InnerText.IndexOf(b));
                            Text t1 = r.Elements<Text>().First();
                            t1.Remove();
                            r.Append(new Text(s1+ textBox2.Text+s2));
                            continue;
                        }
                        string temp = r.InnerText.Substring(0,r.InnerText.IndexOf(a)+a.Length) ;
                        Text text = r.Elements<Text>().First();
                        text.Remove();
                        r.Append(new Text(temp));
                        flag = 1;
                        continue;
                        
                    }
                    if(flag==1)
                    {
                        if(r.InnerText.Contains(b))
                        {
                            string temp = r.InnerText.Substring(r.InnerText.IndexOf(b));
                            Text text = r.Elements<Text>().First();
                            text.Remove();
                            r.Append(new Text(textBox2.Text+temp));
                        }
                        else
                        {
                            list.Add(r);
                            //r.Remove();
                        }
                    }
                }
                foreach(Run rr in list)
                {
                    rr.Remove();
                }
                
                
            }
        }
        void DeleteAlpha(Run run)
        {
            if (run.Elements<Text>().Count()==0) return;
            string strRemoved = Regex.Replace(run.InnerText, "[a - z]", "", RegexOptions.IgnoreCase);
            Text text = run.Elements<Text>().First();
            text.Remove();
            run.Append(new Text(strRemoved));
        }

        void DeleteNumber(Run run)
        {
            if (run.Elements<Text>().Count() == 0) return;
            string strRemoved = Regex.Replace(run.InnerText, @"\d{7,11}$", "", RegexOptions.IgnoreCase);
            Text text = run.Elements<Text>().First();
            text.Remove();
            run.Append(new Text(strRemoved));
        }

        void DeleteLink(Run run)
        {
            Regex r = new Regex("<w:rStyle(.)*?/>");
            //run.InnerXml = r.Replace(run.InnerXml, "");
            if (run.InnerText.Contains("HYPERLINK"))
            {

                //Text text = run.Elements<Text>().First();
                //text.Remove();
                r = new Regex("<w:fldChar(.)*?/>");
                //run.InnerXml = r.Replace(run.InnerXml,"<w:t></w:t>");
                Paragraph pp = (Paragraph)run.Parent;
                run.Remove();
                pp.InnerXml = r.Replace(pp.InnerXml,"");
                r = new Regex("<w:rStyle(.)*?/>");
                pp.InnerXml = r.Replace(pp.InnerXml, "");

            }
        }

        void DeleteBefore(Paragraph p)
        {
            string a = textBox6.Text;
            List<Run> list = new List<Run>();
            if(p.InnerText.Contains(a))
            {
                foreach(Run r in p.Elements<Run>())
                {
                    if(!r.InnerText.Contains(a))
                    {
                        list.Add(r);
                        //r.Remove();
                    }
                    else
                    {
                        string temp = r.InnerText.Substring(r.InnerText.IndexOf(a)+a.Length);
                        Text text = r.Elements<Text>().First();
                        text.Remove();
                        r.Append(new Text(temp));
                        foreach(Run rr in list)
                        {
                            rr.Remove();
                        }
                        return;
                    }
                }
            }
        }

        void DeleteAfter(Paragraph p)
        {
            string a = textBox1.Text;
            int flag = 0;
            List<Run> list = new List<Run>();
            if (p.InnerText.Contains(a))
            {
                foreach (Run r in p.Elements<Run>())
                {
                    if (r.InnerText.Contains(a)&&flag==0)
                    {
                        string temp = r.InnerText.Substring(0,r.InnerText.IndexOf(a));
                        Text text = r.Elements<Text>().First();
                        
                        text.Remove();
                        r.Append(new Text(temp));
                        flag = 1;
                    }
                    if(flag==1)
                    {
                        list.Add(r);
                        //r.Remove();
                    }
                }
            }
            foreach(Run rr in list)
            {
                rr.Remove();
            }
        }







    }

   

}
