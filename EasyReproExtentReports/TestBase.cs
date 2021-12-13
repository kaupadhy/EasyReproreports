using AventStack.ExtentReports;
using AventStack.ExtentReports.MarkupUtils;
using AventStack.ExtentReports.Reporter;
using AventStack.ExtentReports.Reporter.Configuration;
using Microsoft.Dynamics365.UIAutomation.Api.UCI;
using Microsoft.Dynamics365.UIAutomation.Browser;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium;
using System;
using System.IO;
using System.Net;
using System.Net.Mail;
using System.Reflection;
using System.Text.RegularExpressions;

namespace EasyReproExtentReports
{
   
    [TestClass]
    public class TestBase
    {
        protected static ExtentReports Extent;
        protected static ExtentTest TestParent;
        protected static ExtentTest Test;
        protected static string AssemblyName;
        public static string Reportname;
        public static string timestamp;
        public TestContext TestContext { get; set; }
        public static String dir;
        [AssemblyInitialize]
        public static void AssemblyInitialize(TestContext context)
        {
            //AssemblyName = Assembly.GetExecutingAssembly().GetName().Name;
            AssemblyName = Assembly.GetExecutingAssembly().GetName().Name;

            // http://extentreports.com/docs/versions/4/net/
            dir = context.TestDir + "\\";
            const string fileName = "report.html";
            var htmlReporter = new ExtentV3HtmlReporter(dir + fileName);
            htmlReporter.Config.DocumentTitle = $"Test Results: {DateTime.Now:MM/dd/yyyy h:mm tt}";
            htmlReporter.Config.ReportName = context.FullyQualifiedTestClassName;
            htmlReporter.Config.Theme = Theme.Standard;

            // Add any additional contextual information
            Extent = new ExtentReports();
            Extent.AddSystemInfo("Browser", Enum.GetName(typeof(BrowserType), TestSettings.Options.BrowserType));
            Extent.AddSystemInfo("Test User", System.Configuration.ConfigurationManager.AppSettings["OnlineUsername"]);
            Extent.AddSystemInfo("D365 CE Instance", System.Configuration.ConfigurationManager.AppSettings["OnlineCrmUrl"]);
            Extent.AttachReporter(htmlReporter);
            context.AddResultFile(fileName);

            // Create a container for the tests in the class
            TestParent = Extent.CreateTest(context.FullyQualifiedTestClassName);
        }

        [AssemblyCleanup]
        public static void AssemblyCleanup()
        {
            Extent.Flush();
            MailMessage message = new MailMessage();
            SmtpClient smtp = new SmtpClient();
            message.From = new MailAddress("kavita.kavita.upadhyay1@gmail.com");
            message.To.Add(new MailAddress("bhatnagard@ccl.org"));
            message.To.Add(new MailAddress("upadhyayk@ccl.org"));
            message.Subject = "Execution Report";
            message.IsBodyHtml = true; //to make message body as html  
            String d=null;
            d=dir.Replace(@"\\",@"\");
           // Regex.Unescape(dir);
            message.Attachments.Add(new Attachment(dir + "report.html"));
            smtp.Port = 587;
            smtp.Host = "smtp.gmail.com"; //for gmail host  
            smtp.EnableSsl = true;
            smtp.UseDefaultCredentials = false;
            smtp.Credentials = new NetworkCredential("kavita.kavita.upadhyay1@gmail.com", "KUKU123456kuku@");
            smtp.DeliveryMethod = SmtpDeliveryMethod.Network;
            smtp.Send(message);
        }

        [TestInitialize]
        public void TestInitialize()
        {
            // Get unit test description attribute
            var type = Type.GetType($"{TestContext.FullyQualifiedTestClassName}, {AssemblyName}");
            var methodInfo = type?.GetMethod(TestContext.TestName);
            var customAttributes = methodInfo?.GetCustomAttributes(false);
            DescriptionAttribute desc = null;
            if (customAttributes != null)
            {
                foreach (var n in customAttributes)
                {
                    desc = n as DescriptionAttribute;
                    if (desc != null)
                        break;
                }
            }

            // Create individual test under the parent container / class
            Test = TestParent.CreateNode(TestContext.TestName, desc?.Description);
        }

        [TestCleanup]
        public void TestCleanup()
        {
            // Sets individual Extent test result so it reflects correctly in the report
            if (Test.Status == Status.Error)
                return;

            switch (TestContext.CurrentTestOutcome)
            {
                case UnitTestOutcome.Error:
                    Test.Fail("Test Failed - System Error");
                    break;
                case UnitTestOutcome.Passed:
                    Test.Pass("Test Passed");
                    break;
                case UnitTestOutcome.Failed:
                    Test.Fail("Test Failed");
                    break;
                case UnitTestOutcome.Inconclusive:
                    Test.Fail("Test Failed - Inconclusive");
                    break;
                case UnitTestOutcome.Timeout:
                    Test.Fail("Test Failed - Timeout");
                    break;
                case UnitTestOutcome.NotRunnable:
                case UnitTestOutcome.Aborted:
                    Test.Skip("Test Failed - Aborted / Not Runnable");
                    break;
                case UnitTestOutcome.InProgress:
                case UnitTestOutcome.Unknown:
                default:
                    Test.Fail("Test Failed - Unknown");
                    break;
            }
        }

        public void AddScreenShot(Microsoft.Dynamics365.UIAutomation.Api.UCI.WebClient client, string title)
        {
            var filename = Guid.NewGuid();
            var filePath = Path.Combine(TestContext.TestResultsDirectory, $"{filename}.png");
            // Wait for the page to be idle (UCI only)
            TimeSpan ts = new TimeSpan(0, 1, 0);
           // client.Browser.Driver.WaitForTransaction(ts);
           // client.Browser.Driver.WaitForTransaction(TimeSpan.5);
            client.Browser.TakeWindowScreenShot(filePath, ScreenshotImageFormat.Png);
            //Test.Info(addScreenCaptureFromPath("pathToScreenshot"));
            Test.Info(title, MediaEntityBuilder.CreateScreenCaptureFromPath(filePath).Build());
        }

        public void LogExceptionAndFail(Exception e)
        {
            // Formats the exception details to look nice
            var message = e.Message + Environment.NewLine + e.StackTrace.Trim();
            var markup = MarkupHelper.CreateCodeBlock(message);
            Test.Error(markup);
            throw e;
        }
    }
}
