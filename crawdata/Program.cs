// See https://aka.ms/new-console-template for more information
using OfficeOpenXml;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System.Data;
using System.Text;

Console.OutputEncoding = Encoding.UTF8;

IWebDriver driver = new ChromeDriver();

driver.Navigate().GoToUrl("https://howkteam.vn/learn");

var lstKtemExcel = new List<KtemExel>();

var itemsCrawls = driver.FindElements(By.CssSelector(".block.block-link-shadow.block-rounded.ribbon.ribbon-bookmark.ribbon-left.ribbon-success"));

foreach (var item in itemsCrawls)
{
    var newKtemExel = new KtemExel
    {
        Title = item.FindElement(By.CssSelector("a>.block-content.block-content-full.block-sticky-options>h4")).Text,
        Views = item.FindElement(By.CssSelector("a>.block-content.block-content-full.block-content-sm.bg-body-light.text-muted.font-size-sm>div:nth-child(2)>strong")).Text,
        LessonNumbers = item.FindElement(By.CssSelector("a>.block-content.block-content-full.block-content-sm.bg-body-light.text-muted.font-size-sm>.d-inline-block.mr-10>strong")).Text,
        Author = item.FindElement(By.CssSelector(".useravatar-edit-container>a>span")).GetAttribute("innerHTML"),
        Thumbnail = item.FindElement(By.CssSelector(".options-container>img")).GetAttribute("src")
    };
    lstKtemExcel.Add(newKtemExel);
}

/* Export Excel */
var ep = new ExcelPackage();
var wb = ep.Workbook;
var ws = wb.Worksheets.Add("CrawlData");
ws.Cells.Style.Font.Name = "Calibri";
ws.Cells.Style.Font.Size = 15;

ws.Cells[1,1].Value = "Tiêu đề";
ws.Cells[1, 1].Style.Font.Bold = true;
ws.Cells[1,2].Value = "Lượt xem";
ws.Cells[1, 2].Style.Font.Bold = true;
ws.Cells[1,3].Value = "Số bài học";
ws.Cells[1, 3].Style.Font.Bold = true;
ws.Cells[1,4].Value = "Tác giả";
ws.Cells[1, 4].Style.Font.Bold = true;
ws.Cells[1,5].Value = "Thumbnail";
ws.Cells[1, 5].Style.Font.Bold = true;


for (int i = 0; i < lstKtemExcel.Count; i++)
{
    for (int j = 1; j < 6; j++)
    {
        var cell = ws.Cells[i + 2, j];
        switch (j)
        {
            case 1:
                cell.Value = lstKtemExcel[i].Title;
                break;
            case 2:
                cell.Value = lstKtemExcel[i].Views;
                break;
            case 3:
                cell.Value = lstKtemExcel[i].LessonNumbers;
                break;
            case 4:
                cell.Value = lstKtemExcel[i].Author;
                break;
            case 5:
                cell.Value = lstKtemExcel[i].Thumbnail;
                break;
            default:
                break;
        }
    }
}
string handle = Guid.NewGuid().ToString();
Directory.CreateDirectory(string.Format("E:\\MyProject\\howKteam\\crawdata\\ExportFile\\{0}", handle));
var filePath = string.Format("E:\\MyProject\\howKteam\\crawdata\\ExportFile\\{0}\\{1}", handle, "crawdata.xlsx");
ep.SaveAs(new FileInfo(filePath));

Console.WriteLine("Export Success");
Console.ReadLine();


public class KtemExel
{
    public string Title { get; set; }
    public string Views { get; set; }
    public string LessonNumbers { get; set; }
    public string Author { get; set; }
    public string Thumbnail { get; set; }
}