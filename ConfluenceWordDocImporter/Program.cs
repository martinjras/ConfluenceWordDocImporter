using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using CommandLine;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.Remote;

namespace ConfluenceWordDocImporter
{
    class Program
    {
        static void Main(string[] args)
        {
            var options = new Options();

            if (!Parser.Default.ParseArguments(args, options))
            {
                return;
            }

            var files = GetFilesToImport(options);

            if (files.Count == 0)
            {
                Console.WriteLine("No doc or docx files were found at the given path");
                return;
            }

            var driver = GetWebDriver(options);

            driver.Manage().Timeouts().ImplicitlyWait(TimeSpan.FromSeconds(3));

            try
            {
                driver.Navigate().GoToUrl(options.ParentPage);
                driver.FindElement(By.Id("os_username")).Clear();
                driver.FindElement(By.Id("os_username")).SendKeys(options.Username);
                driver.FindElement(By.Id("os_password")).Clear();
                driver.FindElement(By.Id("os_password")).SendKeys(options.Password);
                driver.FindElement(By.Id("loginButton")).Click();

                Thread.Sleep(3000);

                foreach (var file in files)
                {
                    driver.Navigate().GoToUrl(options.ParentPage);
                    driver.FindElement(By.CssSelector("#action-menu-link > span > span")).Click();
                    driver.FindElement(By.Id("import-word-doc")).Click();
                    driver.FindElement(By.Id("filename")).SendKeys(file);
                    driver.FindElement(By.Id("next")).Click();
                    driver.FindElement(By.Name("submit")).Click();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }
            finally
            {
                driver.Close();
            }

        }

        private static RemoteWebDriver GetWebDriver(Options options)
        {
            if(options.Browser == Browser.Chrome)
                return new ChromeDriver();

            return new FirefoxDriver();
        }

        private static List<string> GetFilesToImport(Options options)
        {
            if (Directory.Exists(options.InputFileFolder))
            {
                return Directory.EnumerateFiles(options.InputFileFolder)
                    .Where(file => file.ToLower().EndsWith("doc") || file.ToLower().EndsWith("docx"))
                    .ToList();
            }

            if (File.Exists(options.InputFileFolder))
            {
                if (options.InputFileFolder.ToLower().EndsWith("doc") ||
                    options.InputFileFolder.ToLower().EndsWith("docx"))
                {
                    return new List<string> { options.InputFileFolder };
                }
            }

            return new List<string>();
        }
    }
}
