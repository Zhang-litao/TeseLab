using GeneWebDriver.Helper;
using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GeneWebDriver
{
    public class GeneInfoScraper
    {
        //https://www.ncbi.nlm.nih.gov/gene/
        public (string AlsoKnown, string Summary) ScrapeGeneInfo( string geneId)
        {
            try
            {
                using var driver = WebChromeHelper.StartWebChrome();
                // 设置页面加载的超时时间为30秒（可以根据需要调整）
                driver.Manage().Timeouts().PageLoad = TimeSpan.FromSeconds(30);
                string AlsoKnown = "";
                string Summary = "";


                string GeneUrl = $"https://www.ncbi.nlm.nih.gov/gene/{geneId}";
                try
                {
                    driver.Navigate().GoToUrl(GeneUrl);
                }
                catch (OpenQA.Selenium.WebDriverException ex)
                {
                    Console.WriteLine($"{GeneUrl}：访问网页失败，已记录");
                   
                    // 1. 创建文件路径
                    string folderPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "ExcelFile");
                    string fileName = $"{DateTime.Now.ToString("MM-dd_HH时mm分")}.txt";
                    string filePath = Path.Combine(folderPath, fileName);

                    // 2. 创建文件
                    using (StreamWriter writer = new StreamWriter(filePath))
                    {
                        // 3. 将文本写入文件
                        writer.WriteLine(GeneUrl);
                        // 可以写入更多内容
                    }
                    return ("", "");

                }
                // Visit the webpage
               

                // Find the div element with id "summaryDiv"
                IWebElement summaryDiv = driver.FindElement(By.Id("summaryDiv"));

                // Function to get the text of the first <dd> element after a <dt> element
                string GetDdText(IWebElement parentElement, string dtText)
                {

                    IWebElement? dtElement = parentElement.FindElement(By.XPath($".//dt[text()='{dtText}']"));
                    if (dtElement != null)
                    {
                        IWebElement ddElement = dtElement.FindElement(By.XPath("./following-sibling::dd[1]"));
                        if (ddElement != null)
                        {
                            return ddElement.Text.Trim();
                        }
                    }
                    return null;

                }

                // Get the "Also known as" text
                AlsoKnown = GetDdText(summaryDiv, "Also known as");

                // Get the "Summary" text
                Summary = GetDdText(summaryDiv, "Summary");


                return (AlsoKnown, Summary);
            }
            catch (NoSuchElementException)
            {

                Console.WriteLine("未找到文本dtElement");
            }
            return ("","");
           
        }
    }
}
