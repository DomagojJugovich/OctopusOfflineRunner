using MailKit.Net.Smtp;
using MailKit.Security;
using MimeKit;
using RtfPipe.Model;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Security.Policy;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Windows.Threading;
using System.Xml;
using ZeroDep;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Window;

namespace OfflineDropRunner
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        OfflineDropSettings runnerSett;
        FlowDocument myFlowDoc;
        Paragraph myParagraph;

        string cmdPath;
        string exePath;

        string rtfLog2save;
        string txtLog2save;
        string htmlLog2save;
        private bool thereWereErrors = false;

        public MainWindow()
        {
            InitializeComponent();
            myFlowDoc = new FlowDocument();
            Bold myBold = new Bold(new Run("LOG::::::::::::::::::" + Environment.NewLine + Environment.NewLine));
            myParagraph = new Paragraph();
            myParagraph.Inlines.Add(myBold);
            myFlowDoc.Blocks.Add(myParagraph);
            this.OutText.Document = myFlowDoc;

            if ( !UACHelper.UACHelper.IsAdministrator ) {
                System.Windows.MessageBox.Show("Molim vas pokrenite instalaciju kao korisnik s administratorskim dozvolama !");
                Environment.Exit(-2); 
            }
            if ( !UACHelper.UACHelper.IsElevated ) {
                System.Windows.MessageBox.Show("Molim vas pokrenite instalaciju kao korisnik s administratorskim dozvolama u elevated modu !");
                Environment.Exit(-3); 
            }
            
        }

        private void RunButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                this.RunButton.IsEnabled = false;

                exePath = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
                cmdPath = System.IO.Directory.GetParent(exePath).FullName;
                string[] dirFilesCmd = System.IO.Directory.GetFiles(cmdPath, "*.cmd", SearchOption.TopDirectoryOnly);
                string[] dirFilesPs1 = System.IO.Directory.GetFiles(cmdPath, "*.ps1", SearchOption.TopDirectoryOnly);

                //Sve bi mogli pretpostaviti ali config offline targeta zahtjeva fixne foldere za app i za journal tako da ovjde imamo config koji sadrzi iste postavke !!!!!!!!!!
                string settFileName = System.IO.Path.Combine(exePath,"offlineRunnerSettings.json");
                string jsonString = File.ReadAllText(settFileName);
                runnerSett = (OfflineDropSettings)Json.Deserialize(jsonString, typeof(OfflineDropSettings));

                //izadji iz subfoldera
                System.IO.Directory.SetCurrentDirectory(cmdPath);
                //ubaci passwor  u ps1
                InsertPassword(dirFilesPs1[0], cmdPath);
                //execute
                ExecCmd(dirFilesCmd[0], this.OutText);


            } catch (Exception ex) {

                logError(ex);

            } 
        }

        private void InsertPassword(string file, string cmdPath)
        {
            string[] lines = File.ReadAllLines(file);
            string[] lines2 = new string[lines.Length];

            for(int i = 0; i < lines.Length; i++)
            {
                if (lines[i].Contains("$Password = ")) {
                    lines2[i] = "    $Password = \"" + TxtPassword.Text + "\"";
                    continue;
                }

                lines2[i] = lines[i];
            }

            File.WriteAllLines(file, lines2);
        }

        private void logError(Exception ex)
        {
            try
            {
                appendOutput("ERROR in OfflineRunner :", Brushes.Red, this.OutText, true, true);
                appendOutput(ex.Message, Brushes.Red, this.OutText, true, true);
                appendOutput(ex.StackTrace, Brushes.Black, this.OutText, false, false);
                if (ex.InnerException != null)
                {
                    appendOutput(ex.InnerException.Message, Brushes.Red, this.OutText, true, true);
                    appendOutput(ex.InnerException.StackTrace, Brushes.Black, this.OutText, true, true);
                }
            } catch { }
        }

        private void CheckSuccess(string tentacleJournal)
        {
           bool success = true;

            XmlDocument document = new XmlDocument();
            document.Load(tentacleJournal);


            //deploymentJournal file id deleted evry toime before run so cjeck all nodes
            appendOutput("DeploymentJournal CHECK :", Brushes.Black, this.OutText, false, false);
                    foreach (XmlNode node in document.DocumentElement.ChildNodes)
                    {
                        if (node.Attributes["WasSuccessful"].Value != "True")
                        {
                            appendOutput("ERROR :::: DeploymentJournal WasSuccessful is False - " + node.Attributes["RetentionPolicySet"].Value, Brushes.Red, this.OutText, false, false);
                            success = false;
                        } else {
                            appendOutput("DeploymentJournal STEP: " + node.Attributes["RetentionPolicySet"].Value + " - is OK !", Brushes.Black, this.OutText, false, false);
                        }
                    }

                   //was there Error Output (redirect)
                    if (thereWereErrors) {
                        appendOutput("ERROR :::: thereWereErrors is True" , Brushes.Red, this.OutText, false, false); 
                        success = false; 
                    }

            if (success)
            {
                appendOutput("DEPLOYMENT SUCCESSFULL !!!!!", Brushes.Green, this.OutText, true, true);
            }
            else
            {
                appendOutput("ERROR :::: DEPLOYMENT UN-SUCCESSFULL !!!!!!!!!!!", Brushes.Red, this.OutText, true, true);
            }

        }

        /// <summary>
        /// Executes command
        /// </summary>
        /// <param name="cmd">command to be executed</param>
        /// <param name="output">output which application produced</param>
        /// <param name="transferEnvVars">true - if retain PATH environment variable from executed command</param>
        /// <returns>true if process exited with code 0</returns>
        private void ExecCmd(string cmd, System.Windows.Controls.RichTextBox block)
        {
            thereWereErrors = false;

            ProcessStartInfo processInfo;
            Process process;

            //if (transferEnvVars)
            //    cmd = cmd + " && echo --VARS-- && set";

            appendOutput(cmd, Brushes.Black, OutText, false, false);
            processInfo = new ProcessStartInfo("cmd.exe", "/c " + "\"" + cmd + "\"");
            processInfo.CreateNoWindow = true;
            processInfo.UseShellExecute = false;
            processInfo.RedirectStandardError = true;
            processInfo.RedirectStandardOutput = true;

            process = new Process();
            process.StartInfo = processInfo;
            process.EnableRaisingEvents = true;
            process.OutputDataReceived += (sender, args) => { appendOutput(args.Data, Brushes.Black, this.OutText, false, false); };
            
            process.ErrorDataReceived += (sender, args) => {
                if (args.Data != null)
                {
                    appendOutput(args.Data.Trim(), Brushes.Red, this.OutText, false, false);
                    //args.Data.Trim() jer se u MORHU dogadja da duhovi ubacuju BLANK na std error output pa nema nista crveno a bude thereWereErrors = TRUE iz nicega tj tih duhova
                    thereWereErrors = true;

                    //TODO ovaj hendker se izvudi u nekom drugom threadu te bi trebalo imati svoj try catch jer se dogadja unhandled exception
                }
            };

            process.Exited += new EventHandler(myProcess_Exited);

            //before start clean DeployJournal, so after deploy we will check all nodes for success, without delete is is hard to know which entrys are new
            File.Delete(runnerSett.TentacleJournal);
            process.Start();
            

            process.BeginOutputReadLine();
            process.BeginErrorReadLine();


            //blocks GUI event events above so commented out, we use Exited event
            //process.WaitForExit();

        }

        private void myProcess_Exited(object sender, EventArgs e)
        {
            try { 

                int exitCode = ((Process)sender).ExitCode;

                appendOutput("Process EXIT Code : " + exitCode.ToString(), Brushes.Black, this.OutText, false, false);

                ((Process)sender).Close();


                //check output, journal.xml
                CheckSuccess(runnerSett.TentacleJournal);


                //save log to file and send email
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    try
                    {
                        SaveLog();
                        if (runnerSett.SendEmail) SendEmail(runnerSett.UseOldSmtpClient);
                        //vrati button
                        this.RunButton.IsEnabled = true;

                    }
                    catch (Exception ex) { logError(ex); }
                    
                }), DispatcherPriority.Normal);

            }
            catch (Exception ex)
            {
                logError(ex);
            }
        }

        private void appendOutput(string data, SolidColorBrush color, System.Windows.Controls.RichTextBox tb, bool bold, bool bigger)
        {
            //if (transferenvVars)
            //{
            //    Regex r = new Regex("--VARS--(.*)", RegexOptions.Singleline);
            //    var m = r.Match(output);
            //    if (m.Success)
            //    {
            //        output = r.Replace(output, "");

            //        foreach (Match m2 in new Regex("(.*?)=([^\r]*)", RegexOptions.Multiline).Matches(m.Groups[1].ToString()))
            //        {
            //            String key = m2.Groups[1].Value;
            //            String value = m2.Groups[2].Value;
            //            Environment.SetEnvironmentVariable(key, value);
            //        }
            //    }
            //}


            //gui log
            Dispatcher.BeginInvoke(new Action(() =>
            {


                    var r = new Run(data);
                    r.Foreground = color;
                    if (bold) r.FontWeight = FontWeights.Bold;
                    if (bigger) r.FontSize = r.FontSize + 5;
                    myParagraph.Inlines.Add(r);
                    myParagraph.Inlines.Add(new LineBreak());
                    scrollViewTb.ScrollToEnd();
     


            }), DispatcherPriority.Normal);

            //necemo liniju po liniju nego na EXITED spremiti sve na disk !!!
            ////path of file
            //var path = System.IO.Path.Combine(cmdPath, "OfflineRunnerLog.txt");
            //using (StreamWriter sw = new StreamWriter(path))
            //{
            //    sw.WriteLine(data);
            //}

        }
        public void SendEmail(bool useOldClient)
        {

            string fileName = System.IO.Path.Combine(exePath, "emailSettings.json");
            string jsonString = File.ReadAllText(fileName);
            EmailSettings emailSett = (EmailSettings)Json.Deserialize(jsonString, typeof(EmailSettings));

            if (useOldClient)
            {
                sendWithNETClient(emailSett);
            }
            else
            {
                sendWithMailKit(emailSett);
            }
        }

        private void sendWithMailKit(EmailSettings emailSett)
        {
            var message = new MimeMessage();
            message.From.Add(new MailboxAddress(emailSett.From, emailSett.From));
            foreach (var em in emailSett.To)
            {
                message.To.Add(new MailboxAddress(em, em));
            }

            message.Subject = emailSett.Subject;

            //var bodytext = new TextPart("plain");
            var bodytext = new TextPart("html");
            bodytext.Text = htmlLog2save;
            message.Body = bodytext;


            using (var client = new MailKit.Net.Smtp.SmtpClient())
            {
                client.ServerCertificateValidationCallback = (s, c, h, e) => true;
                client.CheckCertificateRevocation = false;

                if (emailSett.StartTls) { client.Connect(emailSett.Host, emailSett.Port, MailKit.Security.SecureSocketOptions.StartTls); }
                else { client.Connect(emailSett.Host, emailSett.Port, MailKit.Security.SecureSocketOptions.None); }

                client.Authenticate(emailSett.Username, emailSett.Password);

                //var sasl = SaslMechanism.Create("LOGIN", Encoding.UTF8, new NetworkCredential(emailSett.Username, emailSett.Password) );
                //client.Authenticate(sasl);

                client.Send(message);
                client.Disconnect(true);
            }
        }

        private void sendWithNETClient(EmailSettings emailSett)
        {
            var smtpClient = new System.Net.Mail.SmtpClient(emailSett.Host)
            {
                Port = emailSett.Port,
                Credentials = new NetworkCredential(emailSett.Username, emailSett.Password),
                EnableSsl = emailSett.StartTls,
            };

            var mailMessage = new System.Net.Mail.MailMessage
            {
                From = new System.Net.Mail.MailAddress(emailSett.From),
                Subject = emailSett.Subject,
                Body = htmlLog2save,
                IsBodyHtml = true,
            };
            foreach (var em in emailSett.To)
            {
                mailMessage.To.Add(em);
            }

            smtpClient.Send(mailMessage);
            smtpClient.Dispose();

        }

        public class OfflineDropSettings
        {
            private string tentacleApplications;
            private string tentacleJournal;
            private bool sendEmail;
            private bool useOldSmtpClient;

            public string TentacleApplications { get => tentacleApplications; set => tentacleApplications = value; }
            public string TentacleJournal { get => tentacleJournal; set => tentacleJournal = value; }
            public bool SendEmail { get => sendEmail; set => sendEmail = value; }
            public bool UseOldSmtpClient { get => useOldSmtpClient; set => useOldSmtpClient = value; }

            public OfflineDropSettings()  {  }

            //public OfflineDropSettings()
            //{

            //    this.TentacleApplications = "C:\\OctoApp";
            //    this.TentacleJournal = "C:\\OctoApp\\WorkingDir\\DeploymentJournal.xml";

            //}

            //{
            //    "tentacleApplications": "C:\\OctoApp",
            //    "tentacleJournal": "C:\\OctoApp\\WorkingDir\\DeploymentJournal.xml"
            //}




    }

        public class EmailSettings
        {
            private string host;
            private int port;
            private bool startTls;
            private string username;
            private string password;
            private string from;
            private string[] to;
            private string subject;

            public string Host { get => host; set => host = value; }
            public int Port { get => port; set => port = value; }
            public bool StartTls { get => startTls; set => startTls = value; }
            public string Username { get => username; set => username = value; }
            public string Password { get => password; set => password = value; }
            public string From { get => from; set => from = value; }
            public string[] To { get => to; set => to = value; }
            public string Subject { get => subject; set => subject = value; }

            public EmailSettings() { }

            //public EmailSettings()
            //{

            //    this.host = "smtp.office365.com";
            //    this.port = 587;
            //    this.startTls = true;
            //    this.username = "lauscc@laus.hr";
            //    this.password = "kJhfV%45FvcAS!23S";
            //    this.from = "support@laus.hr";
            //    this.to = new string[1];
            //    this.to[0] = "domagoj@laus.hr";
            //    this.to[1] = "domagoj.jugovic@gmail.com";
            //    this.subject = "MORH Offline Drop";

            //}


            //JSON

            //{
            //"host": "smtp.office365.com",
            //"port": 587,
            //"startTls": true,
            //"username": "lauscc@laus.hr",
            //"password": "kJhfV%45FvcAS!23S",
            //"from": "support@laus.hr",
            //"to": [
            //  "domagoj@laus.hr",
            //  "domagoj.jugovic@gmail.com"
            //],
            //"subject": "MORH Offline Drop"
            //}
        }

        private void CopyLog_Click(object sender, RoutedEventArgs e)
        {
            string rtf;
            //for RTF
            using (var memoryStream = new MemoryStream())
            {
                new TextRange(myFlowDoc.ContentStart, myFlowDoc.ContentEnd).Save(memoryStream, System.Windows.DataFormats.Rtf);
                rtf = Encoding.Default.GetString(memoryStream.GetBuffer());
            }

            //for txt
            TextRange textRangeTxt = new TextRange(myFlowDoc.ContentStart, myFlowDoc.ContentEnd);

            if ((bool)this.RtfFormat.IsChecked) {
                System.Windows.Forms.Clipboard.SetText(rtf, (System.Windows.Forms.TextDataFormat)System.Windows.TextDataFormat.Rtf);
            } else {
                System.Windows.Forms.Clipboard.SetText(textRangeTxt.Text, (System.Windows.Forms.TextDataFormat)System.Windows.TextDataFormat.Text);
            }
        }

        private void Log2Mem()
        {

            //for RTF
            using (var memoryStream = new MemoryStream())
            {
                new TextRange(myFlowDoc.ContentStart, myFlowDoc.ContentEnd).Save(memoryStream, DataFormats.Rtf);
                rtfLog2save = Encoding.Default.GetString(memoryStream.GetBuffer());
            }

            //forHtml
            htmlLog2save = RtfPipe.Rtf.ToHtml(rtfLog2save);

            //for txt
            TextRange textRangeTxt = new TextRange(myFlowDoc.ContentStart, myFlowDoc.ContentEnd);
            txtLog2save = textRangeTxt.Text;
                 
        }

        private void SaveLog()
        {
            //posto se saveLog poziva iz EXITED eventa koji je na drugom threadu moracemo i ovdje sa dispatcherom
            var pathTxt = System.IO.Path.Combine(cmdPath, "OfflineRunnerLog.txt");
            var pathRtf = System.IO.Path.Combine(cmdPath, "OfflineRunnerLog.rtf");

            Log2Mem();

            File.WriteAllText(pathTxt, txtLog2save);
            File.WriteAllText(pathRtf, rtfLog2save);


        }
    }
}
