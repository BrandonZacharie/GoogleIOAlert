using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Net;
using System.IO;
using System.Threading;
using Microsoft.Office.Interop.Outlook;
using System.Text.RegularExpressions;
using System.Net.Mail;

namespace GoogleIOAlert
{
    class Program
    {
        private const string url = "http://www.google.com/io";
        private const string urlRedirect = "http://www.google.com/events/io/2011/index-live.html";
        private const string userAgent = "Mozilla/5.0 (Windows NT 6.1) AppleWebKit/535.11 (KHTML, like Gecko) Chrome/17.0.963.56 Safari/535.11";
        private const int minBackoffTime = 10;               //10 seconds
        private const string emailSubject = "Google I/O 2012 Site Updated";
        private const string emailBody = "Quick, go to " + url + "\n Maybe they opened registration!";

        private const string emailRegex = @"^(?("")(""[^""]+?""@)|(([0-9a-z]((\.(?!\.))|[-!#\$%&'\*\+/=\?\^`\{\}\|~\w])*)(?<=[0-9a-z])@))" +
        @"(?(\[)(\[(\d{1,3}\.){3}\d{1,3}\])|(([0-9a-z][-\w]*[0-9a-z]*\.)+[a-z0-9]{2,17}))$";

        private static HttpWebRequest currentRequest;
        private static HttpWebResponse currentResponse;
        private static StringBuilder output;
        private static int backoffTime = minBackoffTime;
        private static string emailAddress;

        private static string sendFromEmailAddress;
        private static string sendFromPassword;

        private enum AlertType
        {
            Outlook = 0,
            Gmail = 1,
            None = 2
        }

        private static string[] alertTypeStrings = { "Outlook", "Gmail", "None" };

        private static AlertType selectedAlertType;
        private static bool validAlertType;
        private static string selectedAlertTypeString;

        static void Main(string[] args)
        {
            Console.Write("How would you like me to alert you? (" + String.Join(", ", alertTypeStrings) + "): ");
            while (!validAlertType)
            {
                selectedAlertTypeString = Console.ReadLine();
                for (int i = 0; i < alertTypeStrings.Length; i++)
                {
                    if (selectedAlertTypeString.Equals(alertTypeStrings[i], StringComparison.OrdinalIgnoreCase))
                    {
                        selectedAlertType = (AlertType)i;
                        validAlertType = true;
                        break;
                    }
                }

                if (validAlertType)
                {
                    break;
                }

                Console.Write("Invalid alert type, try again: ");
            }

            if (selectedAlertType == AlertType.Gmail)
            {
                Console.Write("Enter your gmail address: ");
                sendFromEmailAddress = Console.ReadLine();

                //Validate email address
                while (!Regex.IsMatch(sendFromEmailAddress, emailRegex, RegexOptions.IgnoreCase))
                {
                    //Get user email address
                    Console.Write("Invalid email address, try again: ");
                    sendFromEmailAddress = Console.ReadLine();
                }

                //Gather password
                Console.Write("Enter your gmail password: ");
                sendFromPassword = Console.ReadLine();
            }

            if (selectedAlertType != AlertType.None)
            {
                //Get user email address
                Console.Write("Enter a valid email address to which you would like me to alert: ");
                emailAddress = Console.ReadLine();

                //Validate email address
                while (!Regex.IsMatch(emailAddress, emailRegex, RegexOptions.IgnoreCase))
                {
                    //Get user email address
                    Console.Write("Invalid email address, try again: ");
                    emailAddress = Console.ReadLine();
                }

                Console.Write("An alert will be sent from " + alertTypeStrings[(int)selectedAlertType]);

                if (selectedAlertType == AlertType.Gmail)
                {
                    Console.Write(" using your gmail account: " + sendFromEmailAddress);
                }

                Console.WriteLine("\nAn email will be sent to " + emailAddress + " as soon as the redirect changes for " + url + "\n");
            }

            int attemptNumber = 1;

            //Listen for changes
            while (true)
            {
                //Set request timeout and user agent
                currentRequest = (HttpWebRequest)HttpWebRequest.Create(url);
                currentRequest.Timeout = backoffTime * 1000;
                currentRequest.UserAgent = userAgent;

                output = new StringBuilder();
                output.Append("Attempt #");
                output.Append(attemptNumber++);
                output.Append(":\n");
                output.Append("Fetching Google IO...\n");

                Console.WriteLine(output.ToString());
                output.Clear();

                try
                {
                    currentResponse = (HttpWebResponse)currentRequest.GetResponse();
                    output.Append("Response URI: ");
                    output.Append(currentResponse.ResponseUri.ToString());
                    output.Append("\n");

                    if (!urlRedirect.Equals(currentResponse.ResponseUri.ToString()))
                    {
                        //If the url redirect has changed
                        output.Append("The Google IO 2012 Site is up!\n");
                        alert();

                        Console.WriteLine(output.ToString());
                        output.Clear();

                        break;
                    }
                    else
                    {
                        output.Append("Nothing has changed yet.\n");
                    }

                    currentResponse.Close();

                    output.Append(url);
                    output.Append(" request successful.");

                    if (backoffTime != Math.Max(minBackoffTime, (int)((float)backoffTime * (2.0f / 3.0f))))
                    {
                        //If the backoff time can be reduced
                        output.Append(" Decreasing backoff time from ");
                        output.Append(backoffTime);
                        output.Append(" seconds to ");

                        backoffTime = Math.Max(minBackoffTime, (int)((float)backoffTime * (2.0f / 3.0f)));

                        output.Append(backoffTime);
                        output.Append(" seconds.");
                    }

                    output.Append("\n");
                }
                catch (WebException e)
                {
                    if (e.Status == WebExceptionStatus.Timeout)
                    {
                        //Backoff a bit
                        output.Append(url);
                        output.Append(" timed out. Increasing backoff time from ");
                        output.Append(backoffTime);
                        output.Append(" seconds to ");

                        backoffTime = (int)Math.Ceiling(backoffTime * 1.5);

                        output.Append(backoffTime);
                    }
                    else
                    {
                        output.Append("Web exception: ");
                        output.Append(e.Message);
                    }

                    output.Append("\n");
                }
                finally
                {
                    output.Append("Waiting ");
                    output.Append(backoffTime);
                    output.Append(" seconds until next request.\n\n");

                    Console.Write(output.ToString());
                    output.Clear();
                    Thread.Sleep(backoffTime * 1000);
                }

            }
            Console.WriteLine("Program ended.");
            Console.WriteLine("Press any key to exit.");
            Console.ReadKey();
        }

        private static void alert()
        {
            switch (selectedAlertType)
            {
                case AlertType.Gmail:
                    //@see http://stackoverflow.com/questions/32260/sending-email-in-net-through-gmail
                    var fromAddress = new MailAddress(sendFromEmailAddress, "Google I/O Alert");
                    var toAddress = new MailAddress(emailAddress, "You");

                    var smtp = new SmtpClient
                    {
                        Host = "smtp.gmail.com",
                        Port = 587,
                        EnableSsl = true,
                        DeliveryMethod = SmtpDeliveryMethod.Network,
                        UseDefaultCredentials = false,
                        Credentials = new NetworkCredential(fromAddress.Address, sendFromPassword)
                    };

                    using (var message = new MailMessage(fromAddress, toAddress)
                    {
                        Subject = emailSubject,
                        Body = emailBody
                    })
                    {
                        smtp.Send(message);
                    }
                    break;
                case AlertType.Outlook:
                    Microsoft.Office.Interop.Outlook.Application oApp = new Microsoft.Office.Interop.Outlook.Application();
                    MailItem newEmail = (Microsoft.Office.Interop.Outlook.MailItem)oApp.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olMailItem);
                    newEmail.Recipients.Add(emailAddress);
                    newEmail.Subject = emailSubject;
                    newEmail.Body = emailBody;
                    newEmail.Send();
                    break;
                case AlertType.None:
                default:
                    break;
            }
        }
    }
}