// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

using Microsoft.Dynamics365.UIAutomation.Api.UCI;
using Microsoft.Dynamics365.UIAutomation.Browser;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium;
using System;
using System.Security;
using System;
using System.Net;
using System.Security;
using Microsoft.Dynamics365.UIAutomation.Api.UCI;
using Microsoft.Dynamics365.UIAutomation.Browser;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium;
using AventStack.ExtentReports;
//using NUnit.Framework;
using System.Reflection;

namespace EasyReproExtentReports
{
    [TestClass]
    public class CreateAccount : TestBase
    {
              //  private readonly SecureString _username = System.Configuration.ConfigurationManager.AppSettings["CRMUser"].ToSecureString();
        //private readonly SecureString _password = System.Configuration.ConfigurationManager.AppSettings["CRMPassword"].ToSecureString();
        //private readonly Uri _xrmUri = new Uri(System.Configuration.ConfigurationManager.AppSettings["CRMUrl"].ToString());
      //  protected static ExtentReports Extent;
       // protected static ExtentTest TestParent;
        //protected static ExtentTest Test;
        //protected static string AssemblyName;
        //public Microsoft.VisualStudio.TestTools.UnitTesting.TestContext TestContext { get; set; }

        private static BrowserOptions _options = new BrowserOptions
        {
            BrowserType = BrowserType.Chrome,
            PrivateMode = true,
            FireEvents = false,
            Headless = false,
            UserAgent = false,
        };
        //private readonly SecureString _username = System.Configuration.ConfigurationManager.AppSettings["OnlineUsername"].ToSecureString();
        //private readonly SecureString _password = System.Configuration.ConfigurationManager.AppSettings["OnlinePassword"].ToSecureString();
        //private readonly Uri _xrmUri = new Uri(System.Configuration.ConfigurationManager.AppSettings["OnlineCrmUrl"]);

        [TestMethod]
        [Description("Test should pass")]
        public void CreateAccount_Pass()
        {
            try
            {
               
               
                    var client = new Microsoft.Dynamics365.UIAutomation.Api.UCI.WebClient(_options);
                    //  Test.Info("Log an information message");
                    //  Test.Warning("Log a warning message");
                    using (var xrmApp = new XrmApp(client))
                    {
                        xrmApp.ThinkTime(5000);
                    String d = @"C:\CCL\EasyReproExtentReports-master\TestResults\Deploy_TALTEAM 2021-12-1319_11_29";
                    String i = d.Replace(@"\\", @"\");
                        // var url = System.Configuration.ConfigurationManager.AppSettings["CRMUrl"].ToString();
                        //var user = System.Configuration.ConfigurationManager.AppSettings["CRMUser"].ToString();
                        //var pwd = System.Configuration.ConfigurationManager.AppSettings["CRMPassword"].ToString();
                    String url = "https://cclsandboxqa.crm.dynamics.com/";
                        Console.WriteLine("Read the iguration Data");

                        client.Browser.Driver.Navigate().GoToUrl(url);
                        //AddScreenShot(client, "After login");
                        //      AddScreenShot(client, "After login");
                        Test.Info("Opened home page");
                        AddScreenShot(client, "Navigated to URL");
                        xrmApp.ThinkTime(5000);
                   
                        var travelerInputElement2 = "//input[@id='i0116']";
                        xrmApp.ThinkTime(8000);
                        //  var xrmBrowser = new Microsoft.Dynamics365.UIAutomation.Api.Browser(TestSettings.Options);
                        client.Browser.Driver.FindElement(By.XPath(travelerInputElement2)).SendKeys("upadhyayk@ccl.org");
                       Test.Info("Email enetered");
                       AddScreenShot(client, "Entered email");
                       var Nextbutton = "//input[@id='idSIButton9']";
                        client.Browser.Driver.FindElement(By.XPath(Nextbutton)).Click();
                        xrmApp.ThinkTime(16000);
                        var travelerInputElement = "/html[1]/body[1]/div[2]/div[2]/main[1]/div[2]/div[1]/div[1]/div[2]/form[1]/div[1]/div[3]/div[1]/div[2]/span[1]/input[1]";
                        client.Browser.Driver.FindElement(By.XPath(travelerInputElement)).Clear();
                        xrmApp.ThinkTime(16000);
                        client.Browser.Driver.FindElement(By.XPath(travelerInputElement)).SendKeys("upadhyayk@ccl.org");
                        xrmApp.ThinkTime(16000);
                         Test.Info("Cleared and entered email");
                       AddScreenShot(client, "Cleared and entered emai");
                        client.Browser.Driver.FindElement(By.XPath("//body/div[2]/div[2]/main[1]/div[2]/div[1]/div[1]/div[2]/form[1]/div[2]/input[1]")).Click();
                        xrmApp.ThinkTime(30000);
                        client.Browser.Driver.FindElement(By.XPath("/html[1]/body[1]/div[2]/div[2]/main[1]/div[2]/div[1]/div[1]/div[2]/form[1]/div[1]/div[3]/div[1]/div[2]/span[1]/input[1]")).SendKeys("4Center@2021!");
                        xrmApp.ThinkTime(40000);
                         Test.Info("Entered password");
                        AddScreenShot(client, "Entered password");
                        client.Browser.Driver.FindElement(By.XPath("//body/div[2]/div[2]/main[1]/div[2]/div[1]/div[1]/div[2]/form[1]/div[2]/input[1]")).Click();
                        Test.Info("Push notification sent");
                        AddScreenShot(client, "Push notification sent");
                        xrmApp.ThinkTime(20000);
                        client.Browser.Driver.FindElement(By.XPath("//body/div[2]/div[2]/main[1]/div[2]/div[1]/div[1]/div[2]/form[1]/div[2]/div[1]/div[3]/div[2]/div[2]/a[1]")).Click();
                        xrmApp.ThinkTime(30000);
                        client.Browser.Driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/form[1]/div[1]/div[1]/div[2]/div[1]/div[1]/div[1]/div[1]/div[2]/div[1]/div[3]/div[1]/div[2]/div[1]/div[3]/div[2]/div[1]/div[1]/div[2]/input[1]")).Click();

                        xrmApp.ThinkTime(180000);

                         AddScreenShot(client, "Clicked on stay signed in");
                         xrmApp.ThinkTime(20000);
                    //xrmApp.ThinkTime(5000);

                    xrmApp.Navigation.OpenApp("Dynamics 365 Unified");
                    AddScreenShot(client, "Navigation to App successfull");
                    Console.WriteLine("Navigation to App successfull");

                        xrmApp.ThinkTime(20000);

                        xrmApp.Navigation.OpenSubArea("Marketing", "Leads");
                    AddScreenShot(client, "Navigation to Subarea successfull");

                    Console.WriteLine("Navigation to SubArea successfull");

                        xrmApp.ThinkTime(20000);

                        xrmApp.CommandBar.ClickCommand("New");

                        Console.WriteLine("Clicked on New to create Record");
                    AddScreenShot(client, "Clicked on New to create Record");
                    //  xrmApp.Entity.GetValue("solutionname");

                    xrmApp.ThinkTime(20000);
                        //  xrmApp.Entity.SetValue("subject", "Kavita");
                        xrmApp.Entity.ClearValue("subject");
                        xrmApp.Entity.Save();
                        // client.Browser.Driver.FindElement(By.XPath("/html[1]/body[1]/div[2]/div[1]/div[4]/div[2]/div[1]/div[1]/div[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[1]/div[1]/section[1]/section[1]/div[1]/div[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[1]/div[3]/div[2]/div[1]/div[2]/div[2]/span[2]")).Click();
                        IWebElement m = client.Browser.Driver.FindElement(By.XPath("/html[1]/body[1]/div[2]/div[1]/div[4]/div[2]/div[1]/div[1]/div[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[1]/div[1]/section[1]/section[1]/div[1]/div[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[1]/div[3]/div[2]/div[1]/div[2]/div[2]/span[2]"));
                        String act = m.Text;
                        // System.out.println("Error message is: " + act);
                        //verify error message with Assertion
                        Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual("Solution Name: Required fields must be filled in.", act);
                        /// html[1] / body[1] / div[2] / div[1] / div[4] / div[2] / div[1] / div[1] / div[1] / div[1] / div[1] / div[1] / div[1] / div[2] / div[1] / div[1] / section[1] / section[1] / div[1] / div[1] / div[1] / div[1] / div[1] / div[1] / div[2] / div[1] / div[3] / div[2] / div[1] / div[2] / div[2] / span[2]
                        //xrmApp.Entity.Save();
                        // Microsoft.VisualStudio.TestTools.UnitTesting.Assert.Throws(() => xrmApp.Entity.Save(), "You must provide a value for Demonym.", ExceptionMessageCompareOptions.Exact);
                        // var ex = Microsoft.VisualStudio.TestTools.UnitTesting.Assert.ThrowsException<Exception>(() => xrmApp.Entity.Save());
                        // Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreSame(ex.Message, "Solution Name: Required fields must be filled in.");
                        // Assert.Throws(() => _xrmBrowser.Entity.Save(0), "You must provide a value for Demonym.", ExceptionMessageCompareOptions.Exact);
                        // xrmApp.Entity.SetValue(new LookupItem { Name = "Solution Name", Value = "test test" });
                        // xrmApp.Entity.SetValue("solutionname", "Test solution");
                        // xrmApp.Entity.SetValue("solutioname", "Test solution");
                        //xrmApp.Entity.SetValue("Solution Name", "Test solution");

                        //Console.WriteLine("Set Value in CRM UCI Form successfull");

                        //xrmApp.ThinkTime(40000);
                        //xrmApp.BusinessProcessFlow.SelectStage("Initial Request");
                        //   Xrm.Page.data.entity.getEntityName();

                    /*
                        Console.WriteLine("Records has been Saved successfully");
                    
                        Screenshot ss = ((ITakesScreenshot)client.Browser.Driver).GetScreenshot();
                        ss.SaveAsFile(@"..\TestResultEvidence\SeleniumTestingScreenshot" + ".jpg");
                        Console.WriteLine("Capture Screenshot as Test Result Evidence");
                        //  xrmApp.BusinessProcessFlow.SelectStage("Qualify");
                        // xrmApp.OnlineLogin.Login(_xrmUri, _username, _password);
                        AddScreenShot(client, "After login");

                    xrmApp.Navigation.OpenApp(UCIAppName.Sales);
                    AddScreenShot(client, $"After OpenApp: {UCIAppName.Sales}");

                    xrmApp.Navigation.OpenSubArea("Sales", "Accounts");
                    AddScreenShot(client, "After OpenSubArea: Sales/Accounts");

                    xrmApp.CommandBar.ClickCommand("New");
                    AddScreenShot(client, "After ClickCommand: New");

                    xrmApp.Entity.SetValue("name", TestSettings.GetRandomString(5, 15));
                    AddScreenShot(client, "After SetValue: name");

                    xrmApp.Entity.Save();
                    AddScreenShot(client, "After Save");

                    Assert.IsTrue(true);
                   */
                }
            }
            catch (Exception e)
            {
                LogExceptionAndFail(e);
            }
        }

        [TestMethod]
        [Description("Test should fail with no error")]
        public void JustFail()
        {
            // Example log entries
            Test.Info("Log an information message");
            Test.Warning("Log a warning message");

            // Test code here

            Assert.Fail("Test failed because the passing criteria was not met");
        }

        [TestMethod]
        [Description("Test should fail due to an error")]
        public void CreateAccount_Error()
        {
            try
            {
                var client = new Microsoft.Dynamics365.UIAutomation.Api.UCI.WebClient(TestSettings.Options);
                using (var xrmApp = new XrmApp(client))
                {
                  //  xrmApp.OnlineLogin.Login(_xrmUri, _username, _password);
                    AddScreenShot(client, "After login");

                    xrmApp.Navigation.OpenApp(UCIAppName.Sales);
                    AddScreenShot(client, $"After OpenApp: {UCIAppName.Sales}");

                    xrmApp.Navigation.OpenSubArea("Sales", "Accounts");
                    AddScreenShot(client, "After OpenSubArea: Sales/Accounts");

                    xrmApp.CommandBar.ClickCommand("New");
                    AddScreenShot(client, "After ClickCommand: New");

                    // Field name is incorrect which will cause an exception
                    xrmApp.Entity.SetValue("name344543", TestSettings.GetRandomString(5, 15));
                    AddScreenShot(client, "After SetValue: name");

                    xrmApp.Entity.Save();
                    AddScreenShot(client, "After Save");

                    Assert.IsTrue(true);
                }
            }
            catch (Exception e)
            {
                LogExceptionAndFail(e);
            }
        }
    }
}