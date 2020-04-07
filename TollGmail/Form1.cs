using Limilabs.Client.IMAP;
using Limilabs.Mail;
using Limilabs.Mail.MIME;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace TollGmail
{
    public partial class Form1 : Form
    {
        const string path = "C:\\BaiTap\\";
        public Form1()
        {
            InitializeComponent();
            CreateFolder(path);
            using (StreamReader reader = new StreamReader(path + "user.txt"))
            {
                while (!reader.EndOfStream)
                {
                    var data = reader.ReadLine();
                    var info = data.Split(';');
                    user.Text = info[0];
                    password.Text = info[1];
                    user.Enabled = false;
                    password.Enabled = false;
                    button1.Enabled = false;
                    break;
                }
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string userName = user.Text;
            string passWord = password.Text;

            if (String.IsNullOrEmpty(userName) || String.IsNullOrEmpty(passWord))
            {
                MessageBox.Show("");
                return;
            }

            using (Imap imap = new Imap())
            {
                try
                {
                    imap.ConnectSSL("imap.gmail.com");
                    imap.Login(userName, passWord);
                }
                catch (Exception)
                {
                    MessageBox.Show("Sai mật khẩu hoặc chưa thiết lập gmail");
                    return;
                }

                using (StreamWriter writer = new StreamWriter(path + "user.txt"))
                {
                    string wr = userName + ";" + passWord;
                    writer.Flush();
                    writer.WriteLine(wr);
                }
                MessageBox.Show("Đăng nhập thành công");
                imap.Close();
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            var fromDate = dateTimePicker1.Value;
            var checkSeen = radioButton1.Checked;
            var checkUnSeen = radioButton2.Checked;
            var userName = user.Text;
            var passWord = password.Text;

            var date = fromDate.ToString("MM/dd/yyyy");

            dataGridView1.Rows.Clear();
            using (Imap imap = new Imap())
            {
                imap.ConnectSSL("imap.gmail.com");
                imap.Login(userName, passWord);

                imap.SelectInbox();
                SimpleImapQuery query = new SimpleImapQuery();

                List<long> uids = new List<long>();
                if(checkSeen == false && checkUnSeen == false)
                {
                    uids = imap.Search().Where(
                        Expression.And(
                            Expression.SentSince(DateTime.Parse(date))
                        )
                    );
                }
                if (checkSeen == true)
                {
                    uids = imap.Search().Where(
                        Expression.And(
                            Expression.SentSince(DateTime.Parse(date)),
                            Expression.HasFlag(Flag.Seen)
                        )
                    );
                }
                if (checkUnSeen == true)
                {
                    uids = imap.Search().Where(
                        Expression.And(
                            Expression.SentSince(DateTime.Parse(date)),
                            Expression.HasFlag(Flag.Unseen)
                        )
                    );
                }
                var a = uids.AsEnumerable().Reverse().Take(50);
                List<GmailModel> gmails = new List<GmailModel>();
                var date1 = fromDate.ToString("dd-MM-yyyy");
                CreateFolder(path + date1);
                foreach (long uid in a)
                {

                    var eml = imap.GetMessageByUID(uid);
                    IMail mail = new MailBuilder().CreateFromEml(eml);
                    Console.OutputEncoding = Encoding.Unicode;
                  
                    foreach (MimeData attachment in mail.Attachments)
                    {
                        CreateFolder(path + date1 + "\\" + mail.From[0].Name);
                        attachment.Save(path + date1 + "\\" + mail.From[0].Name + "\\" + attachment.SafeFileName);
                    }
                    dataGridView1.Rows.Add(mail.From[0].Name, mail.Text, mail.Date, mail.Attachments.Count);
                }

                imap.Close();
            }
        }
        private static void CreateFolder(string path)
        {
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }
        }
    }
}
