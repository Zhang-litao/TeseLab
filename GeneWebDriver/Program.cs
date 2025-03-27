
using GeneWebDriver;
using GeneWebDriver.Helper;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using OfficeOpenXml;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using Org.BouncyCastle.Bcpg.Sig;
using System.Collections.Generic;
using System.Security.Cryptography;

//excel 从第row行开始插入
int row = 26830;

//记录当前是在执行基因名称的行号
int CurrentNumber = 222;
//注释
//从第‘CurrentNumber’行开始扫描
List<string> GeneNameValues = ExcelHelper.GetReadData(CurrentNumber);

//页面1
string GeneUrl = "https://www.ncbi.nlm.nih.gov/gene/?term=";
GeneInfoScraper scraper = new GeneInfoScraper();
//表格页数
int TablePageCount = 0;

//临时存储数据
List<GeneTable> geneTables = new List<GeneTable>();

using var driver = WebChromeHelper.StartWebChrome();





foreach (string Gene in GeneNameValues)
{
    try
    {
        Console.WriteLine($"<<<<<<<<<<<<<<<<No.{CurrentNumber}:开始解析查询基因：{Gene}>>>>>>>>>>>>>>>>>");
      
        driver.Manage().Timeouts().PageLoad = TimeSpan.FromSeconds(120); // 增加为60秒，或根据需要调整

        driver.Navigate().GoToUrl(GeneUrl + Gene);



        //// 找到基因搜索框//输入基因信息
        //IWebElement GeneSearch = driver.FindElement(By.Id("term")); GeneSearch.SendKeys(Gene);

        //// 找到搜索按钮并点击
        //IWebElement loginButton = driver.FindElement(By.Id("search")); loginButton.Click();

        //// 等待一些时间以确保登录完成（你可以使用显式等待来实现更精确的等待）
        //System.Threading.Thread.Sleep(5000);

        ////当前网页
        //string CurrentUrl = driver.Url;

        // 查找具有class为"ui-ncbigrid-inner-div"的元素
        IReadOnlyCollection<IWebElement> elements = driver.FindElements(By.ClassName("ui-ncbigrid-inner-div"));
        //是否存在基因列表的表格   没有就跳出

        if (elements.Count <= 0)
        {
            Console.WriteLine($"--------------No.{CurrentNumber}:基因名称：{Gene}未查询到对应的基因列表--------------");
            CurrentNumber++;
            continue;
        }

        // 查找具有class为"title_and_pager"的元素
        IWebElement? pagination = driver.FindElement(By.ClassName("title_and_pager"));

        // 查找具有class为"pagination"的元素
        IWebElement? paginationElement = pagination.FindElement(By.ClassName("pagination"));


        // 在paginationElement下查找第一个h3标签
        IWebElement h3Element = paginationElement.FindElement(By.TagName("h3"));

        // 在h3Element下查找第一个input标签
        IWebElement inputElement = h3Element.FindElement(By.TagName("input"));

        // 获取inputElement的"last"属性的内容
        string lastAttributeValue = inputElement.GetAttribute("last");

        //获取总页数
        TablePageCount = Convert.ToInt32(lastAttributeValue);

        string Url = driver.Url;
        //循环页数
        for (int i = 1; i <= TablePageCount; i++)
        {

            Console.WriteLine($"<<<<<<<<<<<<<<开始编译第：{i}页内容>>>>>>>>>>>>>>>>>>\n\n");
            // 找到页码搜索框并输入页码
            try
            {
                IWebElement page = driver.FindElement(By.Name("EntrezSystem2.PEntrez.Gene.Gene_ResultsPanel.Entrez_Pager.cPage"));
                page.Clear(); // 清除输入框中的内容
                page.SendKeys(i.ToString());
                // 模拟按下Enter键
                page.SendKeys(Keys.Enter);
                // 使用 XPath 定位最后一个 <tbody> 元素
                IWebElement lastTbody = driver.FindElement(By.XPath("(//tbody)[last()]"));

                // 使用 XPath 获取最后一个 <tbody> 元素内的所有 <tr> 标签
                IList<IWebElement> trElements = lastTbody.FindElements(By.TagName("tr"));
                //// 遍历并处理每个 <tr> 元素
                foreach (IWebElement trElement in trElements)
                {


                    IWebElement? value1 = null;
                    IWebElement? value2 = null;
                    IWebElement? value3 = null;
                    try
                    {
                        // 使用 XPath 获取第一个 <td> 标签内的 <span> 标签，并且 class="gene-id"
                        value1 = trElement.FindElement(By.XPath("./td[1]/span[@class='gene-id']"));
                        // 获取 span 标签的文本内容
                        string geneIdText = value1.Text;

                        // 使用 XPath 获取第一个 <td> 标签内的 <span> 标签，并且 class="highlight"
                        value2 = trElement.FindElement(By.XPath("./td[1]/div[2]/a"));
                        string AttendantName = value2.Text;

                        // 使用 XPath 获取第二个 <td> 标签
                        value3 = trElement.FindElement(By.XPath("./td[2]"));
                        string Descripition = value3.Text;
                        if (!Descripition.ToLower().Contains("human")) continue;

                        GeneTable table = new GeneTable()
                        {
                            Gene_ID = geneIdText.Replace("ID:", "").Trim(),
                            AttendantName = AttendantName,
                            Description = Descripition
                        };

                        GetAlsoKnownAndSummary(table, Gene);
                    }
                    catch (OpenQA.Selenium.WebDriverException)
                    {
                        Console.WriteLine($"元素未找到");
                        continue;
                    }

                }
            }
            catch (WebDriverException e)
            {
               
                continue;
            }






            Console.WriteLine($"<<<<<<<<<<<<<<编译完成第{i}页>>>>>>>>>>>>>>>>>>\n\n");
        }

        Console.WriteLine($"<<<<<<<<<<<<<<<<No.{CurrentNumber}:基因解析结束：{Gene}>>>>>>>>>>>>>>>>>");
        CurrentNumber++;
    }
    catch (NoSuchElementException)
    {
        

        Console.WriteLine($"当前基因{Gene}未找到文本当前table页");
        continue;
    }




}





void GetAlsoKnownAndSummary(GeneTable table, string GeneName)
{
    var (AlsoKnown, Summary) = scraper.ScrapeGeneInfo(table.Gene_ID);

    if ((AlsoKnown != "" && AlsoKnown != null) && (Summary != null && Summary != ""))
    {
        Console.WriteLine($"Rows:{CurrentNumber}");
        Console.WriteLine($"GeneName:{GeneName}");
        Console.WriteLine($"Gene ID: {table.Gene_ID}");
        Console.WriteLine($"AttendantName: {table.AttendantName}");
        Console.WriteLine($"Description: {table.Description}");
        Console.WriteLine($"Also Known As: {AlsoKnown}");
        Console.WriteLine($"Summary: {Summary}");
        Console.WriteLine();
        ExportFile(table, GeneName, AlsoKnown, Summary);

    }
    else
    {
        Console.WriteLine($"Gene ID: {table.Gene_ID}");
        Console.WriteLine("Information not found.");
        Console.WriteLine();
    }
}



/// <summary>
/// 筛选出来的基因写入excel
/// </summary>
/// <param name="tables"></param>
void ExportFile(GeneTable tables, string GeneName, string AlsoKnown, string Summary)
{

    string fileurl = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "ExcelFile", "符合条件的基因ID.xlsx");
    //使用EPPlus库创建一个ExcelPackage对象，用于读取或写入Excel文件
    ExcelPackage package = new ExcelPackage(fileurl);
    //设置ExcelPackage的许可证上下文为非商业用途
    ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
    var worksheet = package.Workbook.Worksheets[0]; // 假设模板在第一个工作表

    // 在 Excel 表格中插入数据
    // 假设数据应从第二行开始

    worksheet.Cells[row, 1].Value = CurrentNumber;
    worksheet.Cells[row, 2].Value = GeneName;
    worksheet.Cells[row, 3].Value = tables.Gene_ID;
    worksheet.Cells[row, 4].Value = tables.AttendantName;
    worksheet.Cells[row, 5].Value = tables.Description;
    worksheet.Cells[row, 6].Value = AlsoKnown;
    worksheet.Cells[row, 7].Value = Summary;
    row++;




    // 保存 Excel 文件
    package.Save();




}

// 关闭浏览器窗口
driver.Quit();




Console.ReadKey();