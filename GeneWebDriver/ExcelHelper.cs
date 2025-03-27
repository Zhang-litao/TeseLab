using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;
using OpenQA.Selenium;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
namespace GeneWebDriver
{
    public class ExcelHelper
    {
        /// <summary>
        /// 获取基因名称
        /// </summary>
        /// <returns></returns>
        public static List<string> GetReadData(int Rows)
        {
            string fileUrl = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "ExcelFile", "CKB基因list及爬取需求.xlsx");
            //使用EPPlus库创建一个ExcelPackage对象，用于读取或写入Excel文件
            ExcelPackage package = new ExcelPackage(fileUrl);

            //设置ExcelPackage的许可证上下文为非商业用途
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

            //选择第二个工作表
            var worksheet = package.Workbook.Worksheets[1];


            //获取 Excel 工作表中所有有数据的行数
            int rowCount = worksheet.Dimension.Rows;

            try
            {
                //临时存储数据
                var data = new List<string>();
                for (int i = Rows; i <= rowCount; i++)
                {
                    data.Add(worksheet.Cells[i, 1].Value.ToString()!);
                }
                return data;


            }
            catch (Exception)
            {

                throw;
            }

            return null;
        }



    }
}
