using EASendMail; //add EASendMail namespace (This needs the license code)
using GNAlibrary;
using Microsoft.Win32;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Configuration;
using System.IO;
using System.Globalization;



namespace NoDataAlarm
{
    class Program
    {
        static void Main(string[] args)
        {
            //
            // There are 2 parts to this alarm program
            //
            // Daily report on the state of the system
            // This reflects the state of the system (all settop units) at the time of the email
            // The main thing is that the email comes out once a day regardless of alarm / no alarm state
            // This confirms that the server is operating and that the alarm email check is working
            //
            // Data flow from the Settop units
            // This checks the time of the last data file received for each connected Settop
            // If it is older than the defined time, then the first of 3 alarm emails is sent out
            // This is repeated for 2 more emails
            // The alarm emails then stop and you only detect that the system has a faulty settop by looking at the daily status email.
            // This stiops hundreds of emails being sent
            // 
            // Can monitor up to 5 settops at a time - increase by adding more folders
            // Problem is that all folders are subject to the same data interval 
            // only works if they are running at approximately the same frequency
            // The gka folder is the folder that contains the date sub folders
            //
            // if testEmail "Yes", then a test email is sent to confirm that the email functions.
            //

            // =====[ Library ]======================================

            gnaToolbox gna = new gnaToolbox();

            // =====[ Define variables ]======================================

            string strProjectTitle = System.Configuration.ConfigurationManager.AppSettings["ProjectTitle"];
            string strGKAfolder_1 = System.Configuration.ConfigurationManager.AppSettings["GKAfolder_1"];
            string strGKAfolder_2 = System.Configuration.ConfigurationManager.AppSettings["GKAfolder_2"];
            string strGKAfolder_3 = System.Configuration.ConfigurationManager.AppSettings["GKAfolder_3"];
            string strGKAfolder_4 = System.Configuration.ConfigurationManager.AppSettings["GKAfolder_4"];
            string strGKAfolder_5 = System.Configuration.ConfigurationManager.AppSettings["GKAfolder_5"];
            string strGKAfolder_6 = System.Configuration.ConfigurationManager.AppSettings["GKAfolder_6"];
            string strGKAfolder_7 = System.Configuration.ConfigurationManager.AppSettings["GKAfolder_7"];
            string strGKAfolder_8 = System.Configuration.ConfigurationManager.AppSettings["GKAfolder_8"];
            string strGKAfolder_9 = System.Configuration.ConfigurationManager.AppSettings["GKAfolder_9"];
            string strGKAfolder_10 = System.Configuration.ConfigurationManager.AppSettings["GKAfolder_10"];



            string strSettop_1 = System.Configuration.ConfigurationManager.AppSettings["Settop_1"];
            string strSettop_2 = System.Configuration.ConfigurationManager.AppSettings["Settop_2"];
            string strSettop_3 = System.Configuration.ConfigurationManager.AppSettings["Settop_3"];
            string strSettop_4 = System.Configuration.ConfigurationManager.AppSettings["Settop_4"];
            string strSettop_5 = System.Configuration.ConfigurationManager.AppSettings["Settop_5"];
            string strSettop_6 = System.Configuration.ConfigurationManager.AppSettings["Settop_6"];
            string strSettop_7 = System.Configuration.ConfigurationManager.AppSettings["Settop_7"];
            string strSettop_8 = System.Configuration.ConfigurationManager.AppSettings["Settop_8"];
            string strSettop_9 = System.Configuration.ConfigurationManager.AppSettings["Settop_9"];
            string strSettop_10 = System.Configuration.ConfigurationManager.AppSettings["Settop_10"];

            string strNoDataInterval = System.Configuration.ConfigurationManager.AppSettings["NoDataInterval_hrs"];
            string strEmailLogin = System.Configuration.ConfigurationManager.AppSettings["EmailLogin"];
            string strEmailPassword = System.Configuration.ConfigurationManager.AppSettings["EmailPassword"];
            string strEmailFrom = System.Configuration.ConfigurationManager.AppSettings["EmailFrom"];
            string testEmail = System.Configuration.ConfigurationManager.AppSettings["testEmail"];
            string timeToSendSystemStatusEmail = System.Configuration.ConfigurationManager.AppSettings["SystemStatusEmail"];
            string strNoDataAlarmRecipients = System.Configuration.ConfigurationManager.AppSettings["NoDataAlarmRecipients"];
            string strSystemStatusRecipients = System.Configuration.ConfigurationManager.AppSettings["systemStatusRecipients"];

            string strFrequencyOfDataChecks_minutes = System.Configuration.ConfigurationManager.AppSettings["FrequencyOfDataChecks_minutes"];

            string strMessage = "";

            // Welcome message
            gna.WelcomeMessage("NoDataAlarm 20220808");

            // check that the software license is OK
            string strSendEmail = "No";

            //string strSoftwareKey = gna.checkSoftwareReferenceDate(strProjectTitle, strEmailLogin, strEmailPassword, strSendEmail);
            //Console.WriteLine("Software license: " + strSoftwareKey);

            //if (strSoftwareKey == "expired")
            //{
            //    Console.WriteLine("Software licence has expired");
            //    Console.ReadKey();
            //    goto TheEnd;
            //}

            string[] strFolder = new string[20];
            strFolder[0] = "";
            strFolder[1] = strGKAfolder_1;
            strFolder[2] = strGKAfolder_2;
            strFolder[3] = strGKAfolder_3;
            strFolder[4] = strGKAfolder_4;
            strFolder[5] = strGKAfolder_5;
            strFolder[6] = strGKAfolder_6;
            strFolder[7] = strGKAfolder_7;
            strFolder[8] = strGKAfolder_8;
            strFolder[9] = strGKAfolder_9;
            strFolder[10] = strGKAfolder_10;

            string[] strSettop = new string[20];
            strSettop[0] = "";
            strSettop[1] = strSettop_1;
            strSettop[2] = strSettop_2;
            strSettop[3] = strSettop_3;
            strSettop[4] = strSettop_4;
            strSettop[5] = strSettop_5;
            strSettop[6] = strSettop_6;
            strSettop[7] = strSettop_7;
            strSettop[8] = strSettop_8;
            strSettop[9] = strSettop_9;
            strSettop[10] = strSettop_10;
            Console.WriteLine("No Data Alarm start......");
            Console.WriteLine(" ");
            Console.WriteLine("Checking Alarm Status...");

            // Alarm state email
            for (int i = 1; i < 11; i++)
            {
                string strGKAfolder = strFolder[i];
                string strSettopNo = strSettop[i];
                if (strGKAfolder == "None") goto NextStep;
                Console.WriteLine("Checking alarm status: " + strGKAfolder + strSettopNo);
                strMessage = strProjectTitle + " (" + strSettopNo + ")";
                gna.noDataAlarm(strSettopNo, strMessage, strGKAfolder, strNoDataInterval, strEmailLogin, strEmailPassword, strEmailFrom, strNoDataAlarmRecipients, testEmail);
           
            }

        NextStep:
            Console.WriteLine("");
            Console.WriteLine("Now send System Status email if time is right...");

            // System Status email
            DateTime statusEmailTime = DateTime.ParseExact(timeToSendSystemStatusEmail, "HH:mm", CultureInfo.InvariantCulture);
            DateTime TodayDate = DateTime.UtcNow;

            int iTimeWindow = Convert.ToInt16(strFrequencyOfDataChecks_minutes);
            int minutes = (((TodayDate - statusEmailTime).Hours * 60) + (TodayDate - statusEmailTime).Minutes);

            Console.WriteLine("Frequency of data checks: " + iTimeWindow);
            Console.WriteLine("Minutes to status trigger time: " + minutes);

            if ((minutes >= 0) && (minutes < iTimeWindow))
            {
                Console.WriteLine("Send status email..");

                gna.systemStatusEmail(strProjectTitle, strSettop, strFolder, strEmailLogin, strEmailPassword, strEmailFrom, strSystemStatusRecipients);
            }
            else
            {
                Console.WriteLine("Do not send status email..");
            }

        TheEnd:

            Console.WriteLine("System Status checked...");


        }
    }
}