// See https://aka.ms/new-console-template for more information

using HtmlAgilityPack;
using OfficeOpenXml;
using System.Web;

HttpClient client = new HttpClient();
var response = client.GetAsync("https://howkteam.vn/learn");
response.Wait();
if (response.IsCompleted)
{
    var content = response.Result.Content.ReadAsStringAsync();
    content.Wait();
    HtmlDocument doc = new HtmlDocument();
    doc.LoadHtml(content.Result.ToString());
    var lstKteamCrawlData = new List<KtemExel>();
    var rows = doc.DocumentNode.SelectNodes("//*[@id=\"course-list\"]/div");
    foreach (var row in rows)
    {
        lstKteamCrawlData.Add(new KtemExel
        {
            Title = HttpUtility.HtmlDecode(row.SelectSingleNode(".//div/a[1]/div[1]/h4").InnerHtml),
            Views = HttpUtility.HtmlDecode(row.SelectSingleNode(".//div/a[1]/div[2]/div[2]/strong").InnerHtml),
            LessonNumbers = HttpUtility.HtmlDecode(row.SelectSingleNode(".//div/a[1]/div[2]/div[1]/strong").InnerHtml),
            Author = HttpUtility.HtmlDecode(row.SelectSingleNode(".//div/div[2]/a[1]/span").InnerHtml),
            Thumbnail = HttpUtility.HtmlDecode(row.SelectSingleNode(".//div/div[1]/img[1]").GetAttributeValue("src", "title"))
        });

    }
    ExportExcelCrawlData(lstKteamCrawlData);
    Console.WriteLine("Finished Crawl Data!!!");
    Console.ReadLine();
}
static void ExportExcelCrawlData(List<KtemExel> lstKtemExcel)
{
    /* Export Excel */
    var ep = new ExcelPackage();
    var wb = ep.Workbook;
    var ws = wb.Worksheets.Add("CrawlDataWithRequest");
    ws.Cells.Style.Font.Name = "Calibri";
    ws.Cells.Style.Font.Size = 15;

    ws.Cells[1, 1].Value = "Tiêu đề";
    ws.Cells[1, 1].Style.Font.Bold = true;
    ws.Cells[1, 2].Value = "Lượt xem";
    ws.Cells[1, 2].Style.Font.Bold = true;
    ws.Cells[1, 3].Value = "Số bài học";
    ws.Cells[1, 3].Style.Font.Bold = true;
    ws.Cells[1, 4].Value = "Tác giả";
    ws.Cells[1, 4].Style.Font.Bold = true;
    ws.Cells[1, 5].Value = "Thumbnail";
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
    Directory.CreateDirectory(string.Format("E:\\MyProject\\howKteam\\crawdata\\ExportFileWithRequest\\{0}", handle));
    var filePath = string.Format("E:\\MyProject\\howKteam\\crawdata\\ExportFileWithRequest\\{0}\\{1}", handle, "crawdataWithRequest.xlsx");
    ep.SaveAs(new FileInfo(filePath));
}
public class KtemExel
{
    public string Title { get; set; }
    public string Views { get; set; }
    public string LessonNumbers { get; set; }
    public string Author { get; set; }
    public string Thumbnail { get; set; }
}