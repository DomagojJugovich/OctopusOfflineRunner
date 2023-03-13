using MailKit.Net.Smtp;
using MimeKit;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
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

        public MainWindow()
        {
            InitializeComponent();
            myFlowDoc = new FlowDocument();
            Bold myBold = new Bold(new Run("LOG::::::::::::::::::" + Environment.NewLine + Environment.NewLine));
            myParagraph = new Paragraph();
            myParagraph.Inlines.Add(myBold);
            myFlowDoc.Blocks.Add(myParagraph);
            this.OutText.Document = myFlowDoc;
        }

        private void RunButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                this.RunButton.IsEnabled = false;

                exePath = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
                cmdPath = System.IO.Directory.GetParent(exePath).FullName;
                string[] dirFiles = System.IO.Directory.GetFiles(cmdPath, "*.cmd", SearchOption.TopDirectoryOnly);

                string settFileName = System.IO.Path.Combine(exePath,"offlineRunnerSettings.json");
                string jsonString = File.ReadAllText(settFileName);
                runnerSett = (OfflineDropSettings)Json.Deserialize(jsonString, typeof(OfflineDropSettings));

                //izadji iz subfoldera
                System.IO.Directory.SetCurrentDirectory(cmdPath);
                ExecCmd(dirFiles[0], this.OutText);


            } catch (Exception ex) {

                logError(ex);

            } 
        }

        private void logError(Exception ex)
        {
            try
            {
                appendOutput("ERROR in runner :", Brushes.Red, this.OutText, true, true);
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


            //treba provjeriti samo LAstiChile jer deploymentJournal fajla se nikad NE GAZI, samo se dododaje po jedan node/zapis po pokretanju deploymenta, tako Octopus radi....AAAAAAA

                    if (document.DocumentElement.LastChild.Attributes["WasSuccessful"].Value != "True")
                    {
                        appendOutput("ERROR :::: WasSuccessful is False - " + document.DocumentElement.LastChild.Attributes["RetentionPolicySet"].Value, Brushes.Red, this.OutText, false, false);
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
        private void ExecCmd(string cmd, RichTextBox block)
        {
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
            process.ErrorDataReceived += (sender, args) => { appendOutput(args.Data, Brushes.Red, this.OutText, false, false); };
            process.Exited += new EventHandler(myProcess_Exited);
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
                        SendEmail();
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

        private void appendOutput(string data, SolidColorBrush color, RichTextBox tb, bool bold, bool bigger)
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
                try
                {

                    var r = new Run(data);
                    r.Foreground = color;
                    if (bold) r.FontWeight = FontWeights.Bold;
                    if (bigger) r.FontSize = r.FontSize + 5;
                    myParagraph.Inlines.Add(r);
                    myParagraph.Inlines.Add(new LineBreak());
                    scrollViewTb.ScrollToEnd();
                }
                catch (Exception ex) { logError(ex); }


            }), DispatcherPriority.Normal);

            //necemo liniju po liniju nego na EXITED spremiti sve na disk !!!
            ////path of file
            //var path = System.IO.Path.Combine(cmdPath, "OfflineRunnerLog.txt");
            //using (StreamWriter sw = new StreamWriter(path))
            //{
            //    sw.WriteLine(data);
            //}

        }
        public void SendEmail()
        {
            
            string fileName = System.IO.Path.Combine(exePath, "emailSettings.json");
            string jsonString = File.ReadAllText(fileName);
            EmailSettings emailSett = (EmailSettings)Json.Deserialize(jsonString, typeof(EmailSettings));

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
            

            using (var client = new SmtpClient())
            {
                if (emailSett.StartTls) { client.Connect(emailSett.Host, emailSett.Port, MailKit.Security.SecureSocketOptions.StartTls); }
                else { client.Connect(emailSett.Host, emailSett.Port, MailKit.Security.SecureSocketOptions.None); }
                client.Authenticate(emailSett.Username, emailSett.Password);
                client.Send(message);
                client.Disconnect(true);
            }
        }


        public class OfflineDropSettings
        {
            private string tentacleApplications;
            private string tentacleJournal;

            public string TentacleApplications { get => tentacleApplications; set => tentacleApplications = value; }
            public string TentacleJournal { get => tentacleJournal; set => tentacleJournal = value; }

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
                new TextRange(myFlowDoc.ContentStart, myFlowDoc.ContentEnd).Save(memoryStream, DataFormats.Rtf);
                rtf = Encoding.Default.GetString(memoryStream.GetBuffer());
            }

            //for txt
            TextRange textRangeTxt = new TextRange(myFlowDoc.ContentStart, myFlowDoc.ContentEnd);

            if ((bool)this.RtfFormat.IsChecked) {
                System.Windows.Forms.Clipboard.SetText(rtf, (System.Windows.Forms.TextDataFormat)TextDataFormat.Rtf);
            } else {
                System.Windows.Forms.Clipboard.SetText(textRangeTxt.Text, (System.Windows.Forms.TextDataFormat)TextDataFormat.Text);
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
