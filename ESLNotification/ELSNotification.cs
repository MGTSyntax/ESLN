using System;
using System.ServiceProcess;
using System.Threading.Tasks;
using MySql.Data.MySqlClient;
using System.Timers;
using System.Configuration;
using System.IO;
using System.Net.Mail;
using System.Net;

namespace ESLNotification
{
    public partial class ELSNotification : ServiceBase
    {
        MySqlConnection myCon;
        MySqlCommand myCmd;
        MySqlDataReader myDr;
        private Timer timer;
        private bool isProcessing = false;

        private static string serverValue = ConfigurationManager.AppSettings["ServerValue"];
        private static string userValue = ConfigurationManager.AppSettings["UserValue"];
        private static string passwordValue = ConfigurationManager.AppSettings["PasswordValue"];
        private static string portValue = ConfigurationManager.AppSettings["PortValue"];
        private readonly string[] mibValues = new string[] { ConfigurationManager.AppSettings["MIB1Value"], ConfigurationManager.AppSettings["MIB2Value"], ConfigurationManager.AppSettings["MIB3Value"], ConfigurationManager.AppSettings["MIB4Value"] };
        private readonly string senderEmail = ConfigurationManager.AppSettings["SenderValue"];
        private readonly string senderPassword = ConfigurationManager.AppSettings["SenderPasswordValue"];
        private readonly string smtpServer = ConfigurationManager.AppSettings["SMTPServerValue"];
        private readonly string monthValue = ConfigurationManager.AppSettings["MonthValue"];
        private readonly string ccValue = ConfigurationManager.AppSettings["CCValue"];
        private readonly string bccValue = ConfigurationManager.AppSettings["BCCValue"];

        public ELSNotification()
        {
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {
            StartTimer();
        }

        private void StartTimer()
        {
            WriteLogs("Timer started at: " + DateTime.Now.ToString());

            DateTime now = DateTime.Now;
            DateTime targetTime = now.Date.AddDays(1).AddMinutes(1);
            //DateTime targetTime = now.Date.AddHours(11).AddMinutes(44);
            if (targetTime <= now)
            {
                targetTime = targetTime.AddDays(1);
            }
            double intervalMilliseconds = (targetTime - now).TotalMilliseconds;

            timer?.Dispose();

            timer = new Timer(intervalMilliseconds);
            timer.Elapsed += TimerElapsed;
            timer.Start();
        }

        private void TimerElapsed(object sender, ElapsedEventArgs e)
        {
            if (isProcessing)
            {
                return;
            }

            isProcessing = true;

            try
            {
                myCon = new MySqlConnection($"Server={serverValue};User={userValue};Password={passwordValue};Port={portValue};");
                foreach (string mibValue in mibValues)
                {
                    if (!string.IsNullOrWhiteSpace(mibValue))
                    {
                        ProcessESLN(mibValue);
                    }
                }
            }
            catch (Exception ex)
            {
                WriteLogs($"Error: {ex.Message}");
            }
            finally
            {
                isProcessing = false;
                myCon?.Close();
            }

            WriteLogs("Timer elapsed at: " + DateTime.Now.ToString());

            StartTimer();
        }

        private void ProcessESLN(string compName)
        {
            try
            {
                myCon.Open();

                string query = "SELECT " +
                    "a.ji_empNo as empno, " +
                    "CONCAT(a.ji_lname,CASE WHEN TRIM(a.ji_extname) <> '' THEN CONCAT(' ', a.ji_extname) ELSE '' END, ', ',a.ji_fname,' ',LEFT(a.ji_mname, 1),'.') as empname, " +
                    "c.email_add as empeadd " +
                    $"FROM {compName}.trans_basicinfo a " +
                    $"INNER JOIN {compName}.trans_jobinfo b ON a.ji_empNo = b.ji_empNo " +
                    "AND b.ji_active = 1 AND (b.ji_conEnd <> '' OR b.ji_conEnd is not null) " +
                    "AND b.ji_jobStat <> 'Processing for Clearance' " +
                    $"LEFT JOIN {compName}.trans_emailadd c ON a.ji_empNo = c.ji_empNo " +
                    $"LEFT JOIN {compName}.trans_persinfo d ON a.ji_empNo = d.ji_empNo " +
                    "WHERE " +
                    $"DATE_SUB(STR_TO_DATE(b.ji_conEnd, '%Y-%m-%d'), INTERVAL {monthValue} MONTH) = CURDATE()";

                myCmd = new MySqlCommand(query, myCon);
                
                using (myDr = myCmd.ExecuteReader())
                {
                    while (myDr.Read())
                    {
                        string empno = myDr["empno"].ToString();
                        string empname = myDr["empname"].ToString();
                        string empeadd = myDr["empeadd"].ToString();

                        Task.Run(() => SendEmailWithEmbeddedPhoto(empeadd, empno, empname, compName));
                    }
                }
            }
            catch (Exception ex)
            {
                WriteLogs($"{compName}: {ex.Message}");
            }
            finally
            {
                myCon.Close();
            }
        }

        private void SendEmailWithEmbeddedPhoto(string recipientEmailVal, string empno, string empname, string mibName)
        {
            try
            {
                // Create a new MailMessage
                MailMessage mail = new MailMessage();

                // Set the sender
                mail.From = new MailAddress(senderEmail);

                // Set the recipient
                mail.To.Add(new MailAddress(recipientEmailVal));

                // Add CC addresses
                string[] splitCC = ccValue.Split(new string[] { "," }, StringSplitOptions.RemoveEmptyEntries);
                foreach (string cc in splitCC)
                {
                    mail.CC.Add(new MailAddress(cc));
                }

                // Add BCC address
                mail.Bcc.Add(new MailAddress(bccValue));

                // Set the subject and body of the email
                mail.Subject = $"Your Security License Expires in {monthValue} Months";
                mail.IsBodyHtml = true;

                // Create HTML content with image
                string htmlBody = $@"Hi {empname}!<br><br>
                    <img src=""cid:esl"" alt=""Advertisement Image"">
                    <br><br>
                    Just a quick heads-up: your security license will expire in {monthValue} months from today. 
                    Please remember to renew it before then to keep everything running smoothly. 
                    <br><br>
                    If you have any questions or need help with the renewal process, feel free to reach out to us.<br>
                    <ul>
                        <li>gpzulueta@mibsec.com.ph - Gener P. Zulueta (0928 520 1850)</li>
                        <li>aehenson@mibsec.com.ph - Annie E. Henson (0928 520 1853)</li>
                        <li>mntacata@mibsec.com.ph - Marchael N. Tacata (0936 144 3126)</li>
                    </ul>
                    Best regards,<br><br>
                    M I B";

                // Create an alternative view with HTML content
                AlternateView htmlView = AlternateView.CreateAlternateViewFromString(htmlBody, null, "text/html");

                string imgPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "esl.png");

                // Load the image from the URL
                LinkedResource imageResource = new LinkedResource(imgPath, "image/jpeg");
                imageResource.ContentId = "esl"; // Set a unique identifier
                htmlView.LinkedResources.Add(imageResource);

                // Add the alternative view to the email
                mail.AlternateViews.Add(htmlView);

                // Create a new SmtpClient
                using (SmtpClient smtpClient = new SmtpClient(smtpServer))
                {
                    // Set the SMTP port and enable SSL if needed
                    smtpClient.Port = 587; // For example, use port 587 for Gmail
                    smtpClient.EnableSsl = true;

                    // Set the sender's credentials
                    smtpClient.Credentials = new NetworkCredential(senderEmail, senderPassword);

                    // Send the email
                    smtpClient.Send(mail);
                }

                WriteLogs($"Notification for {empno} - {empname} in {mibName} was sent.");
            }
            catch (Exception ex)
            {
                WriteLogs($"SendEmailWithEmbeddedPhoto - {ex.Message}");
            }
        }

        private void WriteLogs(string message)
        {
            string logPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Logs");
            Directory.CreateDirectory(logPath);

            string logFilePath = Path.Combine(logPath, $"ServiceLog_{DateTime.Now.ToShortDateString().Replace('/', '_')}.txt");

            using (StreamWriter sw = File.AppendText(logFilePath))
            {
                sw.WriteLine(message);
            }
        }

        protected override void OnStop()
        {
            timer.Stop();
            timer.Dispose();
            myCon?.Close();
        }
    }
}
