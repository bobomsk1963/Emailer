using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using System.Net;
using System.Net.Mail;
using System.IO;
using System.Net.Mime;
using System.Xml.Serialization;
using System.Threading;
namespace Email_тест
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        static public void SendMail(SaveMailParam smp, StateMail userState)   //Передать параметром чтобы не вызывал сообщения     
        {
            MailMessage mess = null;
            SmtpClient client = null;
            try
            {
                client = new SmtpClient(smp.SMTPServer, smp.SMTPPort);
                client.Credentials = new NetworkCredential(smp.FromMail.Split('@')[0], smp.FromPassword);
                //Выключаем или включаем SSL - (например для гугла и mail должен быть включен).
                client.EnableSsl = smp.SSL;
                client.DeliveryMethod = SmtpDeliveryMethod.Network;

                mess = new MailMessage();
                mess.From = new MailAddress(smp.FromMail, smp.DisplayName); //отображаемое имя                

                string[] masTo = smp.ToMail.Split(';');
                if (masTo.Length > 0)
                {
                    for (int i = 0; i < masTo.Length; i++)
                    {
                        mess.To.Add(new MailAddress(masTo[i].Trim()));
                    }
                }


                mess.Subject = smp.Tema;
                mess.Body = smp.Text;

                mess.IsBodyHtml = smp.isHTML;                               
                                

                client.SendCompleted += new SendCompletedEventHandler(SendCompletedCallback);
            }
            catch (Exception e)
            {
                if (userState != null)
                {
                    if (userState.MessageBoxShow)
                    {
                        MessageBox.Show(e.Message);
                    }
                    userState.EvenWaitH.Set();
                }
                else 
                {
                    MessageBox.Show(e.Message);
                }
                return;
            }

            if ((smp.FileNames!=null) && (smp.FileNames.Count > 0))
            {
                for (int i = 0; i < smp.FileNames.Count; i++)
                {
                    //Теперь прикрепим файл к сообщению...
                    try
                    {
                        string file = smp.FileNames[i];  // Тип файла не определен
                        Attachment attach = new Attachment(file, MediaTypeNames.Application.Octet);
                        // Добавляем информацию для файла
                        ContentDisposition disposition = attach.ContentDisposition;
                        disposition.CreationDate = System.IO.File.GetCreationTime(file);
                        disposition.ModificationDate = System.IO.File.GetLastWriteTime(file);
                        disposition.ReadDate = System.IO.File.GetLastAccessTime(file);
                        mess.Attachments.Add(attach);
                    }
                    catch (Exception e)
                    {

                        if (userState != null)
                        {
                            if (userState.MessageBoxShow)
                            {
                                MessageBox.Show(smp.FileNames[i] + "  -  " + e.Message);
                            }
                            userState.EvenWaitH.Set();
                        }
                        else 
                        {
                            MessageBox.Show(smp.FileNames[i] + "  -  " + e.Message);
                        }
                        return; 
                    }
                }
            }

            client.SendAsync(mess, userState);

        }


        private void button1_Click(object sender, EventArgs e)
        {


            List<string> lfile = null;
            if (listView1.Items.Count > 0)
            {
                lfile = new List<string>();
                for (int i = 0; i < listView1.Items.Count; i++)
                {
                    lfile.Add(listView1.Items[i].Text);
                }
            }

            SaveMailParam smp = new SaveMailParam(textBox3.Text, Int32.Parse(textBox7.Text), textBox4.Text, textBox5.Text, checkBox1.Checked,
                                                  textBox1.Text, textBox6.Text, checkBox2.Checked, textBox2.Text, lfile,textBox8.Text);

            SendMail(smp, null);

        }

        private static void SendCompletedCallback(object sender, AsyncCompletedEventArgs e)
        {

            if (e.UserState != null )
            {
                if (((StateMail)e.UserState).MessageBoxShow)
                {
                    if (e.Cancelled)
                    {
                        MessageBox.Show("Передача остановлена!");//+ token);
                    }
                    if (e.Error != null)
                    {
                        MessageBox.Show("Ошибка - "/*+ token+" - "*/+ e.Error.ToString());
                    }
                    else
                    {
                        MessageBox.Show("Сообщение передано!");//+token);
                    }
                }

                ((StateMail)e.UserState).Result = e.Error == null;
                // ("Сообщение передано -- !");
                ((StateMail)e.UserState).EvenWaitH.Set();
            }
            else
            {

                if (e.Cancelled)
                {
                    //if (token!=null)
                    MessageBox.Show("Передача остановлена!");//+ token);
                }
                if (e.Error != null)
                {
                    //if (token != null)
                    MessageBox.Show("Ошибка - "/*+ token+" - "*/+ e.Error.ToString());
                }
                else
                {
                    //if (token != null)
                    MessageBox.Show("Сообщение передано!");//+token);
                }

            }

            //mailSent = true;
        }


        bool Serial(string fname, SaveMailParam smp)
        {
            bool ret = false;
            try
            {
                XmlSerializer formatter = new XmlSerializer(typeof(SaveMailParam));
                using (FileStream fs = new FileStream(fname, FileMode.OpenOrCreate))
                {
                    formatter.Serialize(fs, smp);
                }
                ret = true;
            }
            catch (Exception e) 
            { 
                MessageBox.Show(fname + "  -  " + e.Message); 
            }
            return ret;
        }

        static public SaveMailParam DeSerial(string fname, bool MessageBoxShow = true)
        {
            SaveMailParam smp = null;
            try
            {
                XmlSerializer formatter = new XmlSerializer(typeof(SaveMailParam));
                using (FileStream fs = new FileStream(fname, FileMode.OpenOrCreate))
                {
                    smp = (SaveMailParam)formatter.Deserialize(fs);
                }
            }
            catch (Exception e)
            {
                if (MessageBoxShow)
                {
                    MessageBox.Show(fname + "  -  " + e.Message);
                }
            }
            return smp;
        }


        private void button2_Click(object sender, EventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Filter = "(*.xml)|*.xml";
            sfd.FilterIndex = 1;

            if (sfd.ShowDialog() == DialogResult.OK)
            {

                List<string> lfile = null; 
                if (listView1.Items.Count > 0)
                {
                    lfile = new List<string>();
                    for (int i = 0; i < listView1.Items.Count; i++)
                    {
                        lfile.Add(listView1.Items[i].Text);
                    }
                }

                SaveMailParam smp = new SaveMailParam(textBox3.Text, Int32.Parse(textBox7.Text), textBox4.Text, textBox5.Text, checkBox1.Checked,
                                                      textBox1.Text, textBox6.Text, checkBox2.Checked, textBox2.Text, lfile,textBox8.Text);

                Serial(sfd.FileName, smp);

            }

        }

        private void Form1_Shown(object sender, EventArgs e)
        {
            textBox4.Text = Properties.Settings.Default.Adres;
            textBox5.Text = Properties.Settings.Default.Password;
            textBox3.Text = Properties.Settings.Default.SMTPServer;
            textBox7.Text = Properties.Settings.Default.PortServer.ToString();
            checkBox1.Checked = Properties.Settings.Default.SSL;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            OpenFileDialog Open = new OpenFileDialog();
            Open.Filter = "(*.*)|*.*";
            Open.FilterIndex = 1;
            Open.Multiselect = true;
            if (Open.ShowDialog() == DialogResult.OK)
            {
                if (Open.FileNames.Length > 0)
                {
                    for (int i = 0; i < Open.FileNames.Length; i++)
                    {
                        FileInfo fileInf = new FileInfo(Open.FileNames[i]);
                        if (fileInf.Exists)
                        {
                            long l = fileInf.Length;
                            ListViewItem lll = listView1.Items.Add(Open.FileNames[i]);
                            lll.SubItems.Add(l.ToString());
                            lll.Tag = l;
                        }
                    }
                }

                long summ = 0;
                for (int i = 0; i < listView1.Items.Count; i++)
                {
                    summ = summ + (long)listView1.Items[i].Tag;
                }

                label8.Text = summ.ToString();
            }            
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void button6_Click(object sender, EventArgs e)
        {
            OpenFileDialog Open = new OpenFileDialog();
            Open.Filter = "(*.xml)|*.xml";
            Open.FilterIndex = 1;

            if (Open.ShowDialog() == DialogResult.OK)
            {
                SaveMailParam smp = DeSerial(Open.FileName);
                if (smp != null)
                {


                    textBox3.Text = smp.SMTPServer;
                    textBox7.Text = smp.SMTPPort.ToString();
                    textBox4.Text = smp.FromMail;
                    textBox5.Text = smp.FromPassword;
                    checkBox1.Checked = smp.SSL;
                    textBox8.Text = smp.DisplayName;

                    textBox1.Text = smp.ToMail;
                    textBox6.Text = smp.Tema;
                    textBox2.Text = smp.Text.Replace("\n",Environment.NewLine);
                    checkBox2.Checked = smp.isHTML;

                    listView1.Items.Clear();

                    if ((smp.FileNames != null) && (smp.FileNames.Count > 0))
                    {
                        for (int i = 0; i < smp.FileNames.Count; i++)
                        {
                            FileInfo fileInf = new FileInfo(smp.FileNames[i]);
                            if (fileInf.Exists)
                            {
                                long l = fileInf.Length;
                                ListViewItem lll = listView1.Items.Add(smp.FileNames[i]);
                                lll.SubItems.Add(l.ToString());
                                lll.Tag = l;
                            }
                        }

                        long summ = 0;
                        for (int i = 0; i < listView1.Items.Count; i++)
                        {
                            summ = summ + (long)listView1.Items[i].Tag;
                        }
                        label8.Text = summ.ToString();
                    }
                }

            }
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            Properties.Settings.Default.Adres = textBox4.Text;
            Properties.Settings.Default.Password = textBox5.Text;
            Properties.Settings.Default.SMTPServer = textBox3.Text;
            Properties.Settings.Default.PortServer = Int32.Parse(textBox7.Text);
            Properties.Settings.Default.SSL = checkBox1.Checked;
            Properties.Settings.Default.Save();
        }

        private void button7_Click(object sender, EventArgs e)
        {
            FormHelp fh = new FormHelp();
            fh.ShowDialog(this);
            fh.Dispose();
        }

        //smtp.mail.ru
    }

    [Serializable]
    public class SaveMailParam
    {
        public string SMTPServer;
        public int SMTPPort;
        public string FromMail;
        public string FromPassword;
        public bool SSL;
        public string DisplayName;

        public string ToMail;
        public string Tema;
        public bool isHTML;
        public string Text;
        public List<string> FileNames;        

        public SaveMailParam()
        { }

        public SaveMailParam(string smtpserver, int smtpport, string frommail, string frompassword, bool ssl,
                      string tomail, string tema, bool ishtml, string text, List<string> filenames, string displayname)
        {
            SMTPServer = smtpserver;
            SMTPPort = smtpport;
            FromMail = frommail;
            FromPassword = frompassword;
            SSL = ssl;
            DisplayName = displayname;

            ToMail = tomail;
            Tema = tema;
            isHTML = ishtml;
            Text = text;
            FileNames = filenames;            
        }
    }

    public class StateMail
    {
        public bool Result = false;
        public EventWaitHandle EvenWaitH = null;
        public bool MessageBoxShow = true;

        public StateMail(EventWaitHandle evenWaitH,bool messageboxshow = true)
        {
            MessageBoxShow = messageboxshow;
            EvenWaitH = evenWaitH;
        }
    }
}
