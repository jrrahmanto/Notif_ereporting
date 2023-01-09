using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using OpenQA.Selenium.Chrome;
using System.Net;
using System.IO;
using System.Globalization;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using WindowsInput;
using WindowsInput.Native;
using System.Data.OleDb;
using OpenQA.Selenium.Remote;

namespace ConsoleNotifEreporting
{
    public class SuiteTests : IDisposable
    {
        public IWebDriver driver { get; private set; }
        public IDictionary<String, Object> vars { get; private set; }
        public IJavaScriptExecutor js { get; private set; }
        static void Main(string[] args)
        {
            var driver = new RemoteWebDriver(new Uri("http://localhost:4444/wd/hub"), new ChromeOptions().ToCapabilities());
            var js = (IJavaScriptExecutor)driver;
            var vars = new Dictionary<String, Object>();
        }

        public SuiteTests()
        {
            driver = new RemoteWebDriver(new Uri("http://localhost:4444/wd/hub"), new ChromeOptions().ToCapabilities());
            js = (IJavaScriptExecutor)driver;
            vars = new Dictionary<String, Object>();
        }
        public void Dispose()
        {
            driver.Quit();
        }
        
        public void Untitled()
        {
            // Test name: Untitled
            // Step # | name | target | value
            // 1 | open | / | 
            driver.Navigate().GoToUrl("https://www.google.com/");
            // 2 | setWindowSize | 1536x824 | 
            driver.Manage().Window.Size = new System.Drawing.Size(1536, 824);
            // 3 | type | name=q | halloo
            driver.FindElement(By.Name("q")).SendKeys("halloo");
            // 4 | sendKeys | name=q | ${KEY_ENTER}
            driver.FindElement(By.Name("q")).SendKeys(OpenQA.Selenium.Keys.Enter);
            // 5 | close |  | 
            driver.Close();
        }
    }
}
