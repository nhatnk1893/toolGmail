using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Limilabs.Client.IMAP;
using Limilabs.Mail;
using Limilabs.Mail.MIME;

namespace Gmail
{
    class Program
    {
        // Generate an unqiue email file name based on date time

        static void Main(string[] args)
        {
            using (Imap imap = new Imap())
            {
                imap.ConnectSSL("imap.gmail.com");
                imap.Login("nhatnk01081993@gmail.com", "niemtin0108");

                imap.SelectInbox();
                SimpleImapQuery query = new SimpleImapQuery();

                //query.Subject = "test";
                //query.From = "nguyenkhanhnhat0108@gmail.com";
                //query.From = "GitHub";
                List<long> uids = imap.Search(Flag.Unseen);
                var a = uids.AsEnumerable().Reverse().Take(50);
                int i = 0;
                foreach (long uid in a)
                {

                    var eml = imap.GetMessageByUID(uid);
                    IMail mail = new MailBuilder().CreateFromEml(eml);
                    Console.OutputEncoding = Encoding.Unicode;
                    Console.WriteLine(mail.Subject);
                    Console.WriteLine(mail.Date);

                    Console.WriteLine(mail.Attachments);
                    Console.WriteLine(mail.From);
                    foreach (MimeData attachment in mail.Attachments)
                    {
                        System.IO.Directory.CreateDirectory(@"c:\folder");
                        attachment.Save(@"c:\folder\" + attachment.SafeFileName);
                    }
                    Console.WriteLine("*********************");
                }
                imap.Close();
            }
        }
    }
}