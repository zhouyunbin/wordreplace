using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Aspose.Words;
using System.IO;
using System.Threading;
using System.Reflection;
using System.Text.RegularExpressions;

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
            radioButton1.Checked = true;

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
        void SaveDocx(string filename)
        {
            Document doc = new Document(filename);
            doc.Save(filename+"x",SaveFormat.Docx);
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
                var doc = new Document(filename.ToString());
                
                DocumentBuilder db = new DocumentBuilder(doc);
                db.MoveToDocumentEnd();
                db.Write("\u0020\u0020");
                doc.Save(filename.ToString());
                doc = new Document(filename.ToString());
                Regex regex1 = new Regex("@|&|\u3000|\u0020|\n");
                doc.Range.Replace(regex1, String.Empty);
                //Regex regex2 = new Regex("\r\r");
                //doc.Range.Replace(regex2, String.Empty);
                if (radioButton1.Checked)
                    ReplaceAandB(doc);
                else if (radioButton2.Checked)
                    ReplaceAB(doc);
                Regex regex = new Regex("\v");
                doc.Range.Replace(regex, new MyReplaceEvaluator(), true);                
                         
                NodeCollection nc;
                //char[] values = doc.GetText().ToCharArray();
                

                //DocumentBuilder builder = new DocumentBuilder(doc);
                
                nc = doc.GetChildNodes(NodeType.Paragraph, true);
                foreach (Paragraph r in nc)
                {
                    string txt = r.GetText();
                    if (txt.Length > 0)
                    {

                        //r.ParagraphFormat.ClearFormatting();
                        r.ParagraphFormat.FirstLineIndent = 0;
                        r.Range.Text.Trim();
                        
                    }
                }

                nc = doc.GetChildNodes(NodeType.Run,true);
                foreach(Run r in nc)
                {
                    string txt = r.GetText();
                    if(r.ParentParagraph.ParagraphFormat.Alignment == ParagraphAlignment.Center)
                    {
                        regex1 = new Regex(r.ParentParagraph.GetText().Trim());
                        r.ParentParagraph.NextSibling.Range.Replace(regex1,String.Empty);
                        continue;
                    }
                    if (txt.Length > 20)
                    {
                        
                       
                        r.ParentParagraph.ParagraphFormat.ClearFormatting();
                        r.ParentParagraph.ParagraphFormat.FirstLineIndent = 0;
                        r.Text=r.Range.Text.TrimStart();
                        if (r.PreviousSibling == null)
                            continue;
                        if (!r.PreviousSibling.GetText().EndsWith("\r"))
                            continue;
                        r.Text = r.Text.Insert(0, "\u3000\u3000");
                        
                    }
                }

                //TwoBlank(doc);
                
                //doc.Range.Replace("\v", "\r", true, true);
                //doc.Replace("", "kkkk", false, false);
                doc.Save(filename.ToString());

                
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
        }

        void ReplaceAandB(Document doc)
        {
            if (textBox1.Text.Trim() != "")
            {
                string orig = textBox1.Text;
                string repl = textBox2.Text;
                string[] orig1 = orig.Split('|');
                string[] repl1 = repl.Split('|');
                int i = 0;
                for (; i < orig1.Length; i++)
                {
                    doc.Range.Replace(orig1[i], repl1[i], true, true);
                }
            }
        }

        void ReplaceAB(Document doc)
        {
            if (textBox1.Text.Trim() != "")
            {
                string orig = textBox1.Text;
                string repl = textBox2.Text;
                string[] orig1 = orig.Split('|');

                Regex regex = new Regex(orig1[0]+".+"+orig1[1]);
                doc.Range.Replace(regex, orig1[0] + repl + orig1[1]);
            }
        }

        void TwoBlank(Document doc)
        {
            Regex regex = new Regex("\r(.{10,}?)\r");
            doc.Range.Replace(regex, new MyReplaceEvaluator2(),true);
        }

        private void button5_Click(object sender, EventArgs e)
        {
            DirectoryInfo theFolder = new DirectoryInfo(textBox3.Text);
            progressBar1.Maximum = theFolder.GetFiles("*.doc*").Length;
            progressBar1.Value = 0;
            Thread th = new Thread(MoveNoTitle);
            th.Start();
        }

        void MoveNoTitle()
        {
            if (!Directory.Exists(textBox3.Text+@"\notitle"))//如果不存在就创建file文件夹
            {
                Directory.CreateDirectory(textBox3.Text + @"\notitle");
            }
            DirectoryInfo theFolder = new DirectoryInfo(textBox3.Text);
            
            foreach (FileInfo NextFile in theFolder.GetFiles())
            {
                if (NextFile.Extension.Equals(".doc") || NextFile.Extension.Equals(".docx"))
                {
                    m_SyncContext.Post(ProgressPlus, null);
                    if (NextFile.Name.Contains("$"))
                    {
                        continue;
                    }
                    m_SyncContext.Post(SetTextSafePost, NextFile.Name);
                    Document doc = new Document(NextFile.FullName);
                    NodeCollection nc = doc.GetChildNodes(NodeType.Run, true);
                    int i = 0;
                    foreach (Run r in nc)
                    {
                        i++;
                        if(i>10)
                        {
                            File.Move(NextFile.FullName, NextFile.Directory.FullName + @"\\notitle\\" + NextFile.Name);
                            break;
                        }
                        string txt = r.GetText();
                        if (r.ParentParagraph.ParagraphFormat.Alignment == ParagraphAlignment.Center)
                        {
                            break;
                        }
                    }

                }
            }

        }
    }

    class MyReplaceEvaluator : IReplacingCallback
    {
        /// <summary>
        /// This is called during a replace operation each time a match is found.
        /// This method appends a number to the match string and returns it as a replacement string.
        /// </summary>
        ReplaceAction IReplacingCallback.Replacing(ReplacingArgs e)
        {
            e.Replacement = "\r";
            
            return ReplaceAction.Replace;
        }
    }

    class MyReplaceEvaluator2 : IReplacingCallback
    {
        /// <summary>
        /// This is called during a replace operation each time a match is found.
        /// This method appends a number to the match string and returns it as a replacement string.
        /// </summary>
        ReplaceAction IReplacingCallback.Replacing(ReplacingArgs e)
        {

            e.Replacement = "\r";
            return ReplaceAction.Replace;
        }
    }

}
