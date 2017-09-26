using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Reflection;
using System.Text;
using System.Threading;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Internal;
using OpenQA.Selenium.Remote;

namespace ProductionMetrics
{
    static class Program
    {
        private static readonly DateTime Date = DateTime.Now.AddDays(-1);
        private static string _reportPath = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location) + @"\Метрики производства.csv";

        static void Main()
        {
            var config = ConfigurationManager.OpenExeConfiguration(Assembly.GetEntryAssembly().Location);
            if (config.AppSettings.Settings.AllKeys.Contains("ReportDestinationPath") &&
                config.AppSettings.Settings["ReportDestinationPath"].Value != "")
            {
                _reportPath = config.AppSettings.Settings["ReportDestinationPath"].Value;
            }
            else
            {
                config.AppSettings.Settings.Remove("ReportDestinationPath");
                config.AppSettings.Settings.Add("ReportDestinationPath", $"{_reportPath = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location) + @"\Метрики производства.csv"}");
                config.Save(ConfigurationSaveMode.Minimal);
            }

            var csvExport = new List<string>();
            if (!File.Exists(_reportPath))
            {
                csvExport.Add(
                    "\"Дата\";" +
                    "\"Входящие инциденты\";" +
                    "\"Обработанные инцидиенты\";" +
                    "\"Открытые баттлы\";" +
                    "\"Решенные баттлы\";" +
                    "\"Переоткрытых баттлы\"");
            }

            var driver = new ChromeDriver();
            var поДатеПеревода = FindAmount(driver,
                "http://wicstat/ReportServer/Pages/ReportViewer.aspx?%2FWic2Reports%2F%D0%AD%D0%BA%D1%81%D0%BF%D0%B5%D1%80%D1%82%D0%BD%D1%8B%D0%B9%20%D0%BE%D1%82%D0%B4%D0%B5%D0%BB%2F%D0%93%D1%80%D1%83%D0%BF%D0%BF%D0%BE%D0%B2%D1%8B%D0%B5%20%D1%82%D0%B5%D0%B3%D0%B8%20%D0%AD%D0%9E%20%D0%BF%D0%BE%20%D0%B4%D0%B0%D1%82%D0%B5%20%D0%BF%D0%B5%D1%80%D0%B5%D0%B2%D0%BE%D0%B4%D0%B0&rc:showbackbutton=true");
            var групповыеТегиПоЭо = FindAmount(driver,
                "http://wicstat/ReportServer/Pages/ReportViewer.aspx?%2FWic2Reports%2F%D0%AD%D0%BA%D1%81%D0%BF%D0%B5%D1%80%D1%82%D0%BD%D1%8B%D0%B9%20%D0%BE%D1%82%D0%B4%D0%B5%D0%BB%2F%D0%AD%D0%BA%D1%81%D0%BF%D0%B5%D1%80%D1%82%D0%BD%D1%8B%D0%B9%20%D0%BE%D1%82%D0%B4%D0%B5%D0%BB.%20%D0%93%D1%80%D1%83%D0%BF%D0%BF%D0%BE%D0%B2%D1%8B%D0%B5%20%D1%82%D0%B5%D0%B3%D0%B8&rc:showbackbutton=true");
            driver.Close();
            driver.Quit();

            var createdBattles = GetYoutrackBillyBattleCount($"created: {Date:yyyy-MM-dd}");
            var resolvedBattles = GetYoutrackBillyBattleCount($"resolved date: {Date:yyyy-MM-dd}");
            var reopenedBattles = GetYoutrackBillyBattleCount($"Регулярность: Периодически,Постоянно updated: {Date:yyyy-MM-dd}");

            csvExport.Add(
                $"\"{Date.ToShortDateString()}\";" +
                $"\"{поДатеПеревода}\";" +
                $"\"{групповыеТегиПоЭо}\";" +
                $"\"{createdBattles}\";" +
                $"\"{resolvedBattles}\";" +
                $"\"{reopenedBattles}\"");
            File.AppendAllLines(_reportPath, csvExport, Encoding.UTF8);
        }

        private static string GetYoutrackBillyBattleCount(string filter)
        {
            using (var httpClient = new HttpClient(new HttpClientHandler()))
            {
                httpClient.DefaultRequestHeaders.Accept.Clear();
                httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer",
                    "perm:Z3JpYg==.WW91VHJhY2tTZWFyY2g=.bSRI8KWwnZX4NPgi4qIYzpX9ib8KlC");
                httpClient.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                string count;
                do
                {
                    count = httpClient
                        .GetAsync(
                            $"https://yt.skbkontur.ru/rest/issue/count?filter=project: Billy Type: Battle {filter}")
                        .Result
                        .Content
                        .ReadAsAsync<IssueCounter>()
                        .Result.Value;
                } while (count == "-1");
                return count;
            }
        }

        private static string FindAmount(RemoteWebDriver driver, string url)
        {
            for (var i = 0; i < 2; i++)
            {
                try
                {
                    driver.Navigate().GoToUrl(url);
                    WaitForRefreshTable(driver);

                    var start = "start";
                    if (driver.FindElementsByCssSelector("[data-parametername='BeginDate']").Count > 0)
                    {
                        start = "BeginDate";
                    }
                    var beginDateInput = driver.FindElementByCssSelector($"[data-parametername='{start}']")
                        .FindElement(By.Id("ReportViewerControl_ctl04_ctl03_txtValue"));
                    beginDateInput.Clear();
                    Thread.Sleep(1500);
                    beginDateInput.SendKeys(Date.ToShortDateString() + Keys.Tab);
                    WaitForRefreshTable(driver);

                    var end = "end";
                    if (driver.FindElementsByCssSelector("[data-parametername='EndDate']").Count > 0)
                    {
                        end = "EndDate";
                    }
                    var endDateInput = driver.FindElementByCssSelector($"[data-parametername='{end}']")
                        .FindElement(By.Id("ReportViewerControl_ctl04_ctl05_txtValue"));
                    endDateInput.Clear();
                    Thread.Sleep(1500);
                    endDateInput.SendKeys(Date.ToShortDateString() + Keys.Tab);
                    WaitForRefreshTable(driver);

                    var searchButton = driver.FindElementByCssSelector("[id='ReportViewerControl_ctl04_ctl00']");
                    searchButton.Click();

                    WaitForRefreshTable(driver);

                    if (driver.FindElementById("ReportViewerControl_ctl05_ctl00_Last_ctl00_ctl00").Displayed)
                    {
                        driver.FindElementById("ReportViewerControl_ctl05_ctl00_Last_ctl00_ctl00").Click();
                    }
                    WaitForRefreshTable(driver);
                    return driver.FindElementByCssSelector("[id$='79iT0_aria']").Text;
                }
                catch (Exception e)
                {
                    Console.WriteLine(e);
                }
            }
            return "ERROR";
        }

        private static void WaitForRefreshTable(IFindsById driver)
        {
            while (driver.FindElementById("ReportViewerControl_AsyncWait_Wait").Displayed)
            {
                Thread.Sleep(400);
            }
            Thread.Sleep(800);
        }
    }
    public class IssueCounter
    {
        public string Value;
    }
}
