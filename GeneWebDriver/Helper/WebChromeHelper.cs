using OpenQA.Selenium.Chrome;
using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using GeneWebDriver.Extensions;

namespace GeneWebDriver.Helper
{
    public class WebChromeHelper
    {
        private static string? _logLocation;

        private static string GetLogLocation()
        {
            if (_logLocation == null || !File.Exists(_logLocation)) _logLocation = Path.GetTempFileName();
            return _logLocation;
        }
        /// <summary>
        /// 加载 Web Chrome
        /// </summary>
        /// <returns>IWebDriver.</returns>
        public static IWebDriver StartWebChrome()
        {
            // cd C:\Program Files\Google\Chrome\Application
            // chrome.exe --remote-debugging-port=9222 --user-data-dir="C:\selenum\AutomationProfile"

            ChromeOptions options = new ChromeOptions();
            options.AddArgument("--incognito");
            options.AddArgument("headless");
            options.AddArgument("Referer=https://www.ncbi.nlm.nih.gov/gene/"); /*116.0.0.0*/
            options.AddArgument("User-Agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/117.0.5938.150 Safari/537.36");
            options.AddArgument("disable-infobars");

            // 浏览器窗口最大化
            options.AddArgument("--start-maximized");

            // 禁用 Chrome 浏览器的自动化信息栏
            options.AddArgument("--disable-infobars");

            // 不显示【正受自动测试软件控制】字样
            options.AddExcludedArguments("enable-automation");

            // 禁用扩展
            options.AddArgument("--disable-extensions");

            // 设置文件下载路径
            options.AddUserProfilePreference("download.default_directory", Path.Combine(Environment.CurrentDirectory, "Down").FolderExists());

            //options.DebuggerAddress = "127.0.0.1:9222";

            // 全屏（但按F11无法退出全屏）
            //options.AddArgument("--kiosk");

            // 全部（按F11可以退出全屏）
            //options.AddArgument("--start-fullscreen");

            //options.AddArguments("--test-type", "--ignore-certificate-errors");

            // 不自动关闭浏览器
            options.LeaveBrowserRunning = true;

            ChromeDriverService chromeDriverService = ChromeDriverService.CreateDefaultService(@"C:\Program Files\Google\Chrome\Application\");
            // 关闭每次调试时打开的CMD
            chromeDriverService.HideCommandPromptWindow = true;
            chromeDriverService.LogPath = GetLogLocation();
            chromeDriverService.EnableVerboseLogging = true;
            chromeDriverService.DisableBuildCheck = true;

            IWebDriver driver = new ChromeDriver(chromeDriverService, options);
            return driver;
        }
    }
}
