using HtmlAgilityPack;
using NPOI;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

const string basePage = @"https://clubeconomy.com.mk";
Console.Write("Enter a valid file path [ex: c:/users/test/desktop/test.xlsx]:");

string filePath = Console.ReadLine().Trim().Replace("/", @"\");

ConsoleKey response;
do
{
    Console.WriteLine("Only search local companies?");
    response = Console.ReadKey(false).Key;
    if(response != ConsoleKey.Enter)
        Console.WriteLine();

} while(response != ConsoleKey.Y && response != ConsoleKey.N);

bool onlyLocalCompanies = false;

onlyLocalCompanies = response == ConsoleKey.Y;

int pageNumber = 1;
bool nextExists = true;

List<List<string>> pageLists = new List<List<string>>();

while (nextExists | pageNumber < 100)
{
    HtmlWeb web = new HtmlWeb();
    var companiesListPage =
            web
            .Load(basePage + "/BusinessAddressBook?pagenumber=" + pageNumber.ToString());

    nextExists = companiesListPage?.DocumentNode.SelectSingleNode(@"//div[@class='no-result']") == null;

    if (nextExists)
    {
        Console.WriteLine($"Processing page: {pageNumber}");
        ProcessPage(companiesListPage?.DocumentNode.SelectNodes("//h2/a[@target='_blank']"), ref pageLists, onlyLocalCompanies);
        Console.WriteLine($"Finished page: {pageNumber}");
    }
    pageNumber++;
}

WriteToExcel(filePath, ref pageLists);

static void WriteToExcel(string filePath, ref List<List<string>> pageLists)
{
    using (var fs = new FileStream(filePath, FileMode.Create, FileAccess.Write))
    {
        IWorkbook workbook = new XSSFWorkbook();
        ISheet worksheet = workbook.CreateSheet("Sheet1");

        for (int i = 0; i < pageLists.Count; i++)
        {
            var row = worksheet.CreateRow(i);
            for (int j = 0; j < pageLists[i].Count; j++)
            {
                row.CreateCell(j).SetCellValue(pageLists[i][j].ToString());
            }
        }
        NormalizeColumnSize(worksheet);
        // Save to file
        workbook.Write(fs);
    }
}

static void ProcessPage(HtmlNodeCollection? nodes, ref List<List<string>> pageLists, bool onlyLocalCompanies)
{
    try
    {
        if (nodes == null || nodes.Count == 0)
            return;

        foreach (var node in nodes)
        {
            var companyTitle = node.InnerText;
            var url = node.Attributes["href"].Value;
            HtmlWeb web = new HtmlWeb();
            var businessPage = web.Load(basePage + url);

            var businessPageNodes =  businessPage.DocumentNode.SelectNodes(@"//div[@class='media-body']//a[@target='_blank']");

            bool filtersFail = false;

            if(onlyLocalCompanies)
                filtersFail = !IsLocalCompany(businessPage.DocumentNode);

            if (businessPageNodes == null || filtersFail)
            {
                continue;
            }

            var companyAddress = GetCompanyAddress(businessPage.DocumentNode);


            var singleCompanyList = new List<string> { companyTitle, companyAddress };
            foreach (var n in businessPageNodes)
            {
                var value = n.Attributes["href"].Value
                    .Trim()
                    .Replace("mailto:", "")
                    .Replace("tel:", "");

                singleCompanyList.Add(value);
            }
            pageLists.Add(singleCompanyList);
        }
    }
    catch (Exception e)
    {
        Console.WriteLine(e.Message);
    }
}

static string GetCompanyAddress(HtmlNode node)
{
    return string.Join(' ', node.SelectNodes(@"//div[@class='media-body'][1]//strong//following-sibling::text()[position() < last()]")
                .Select(x => x.InnerText).ToList());
            
}

static bool IsLocalCompany(HtmlNode node)
{
    var cityName = node.SelectSingleNode(@"//div[@class='media-body'][1]//strong//following-sibling::text()[2]").InnerText;

    if(cityName.Trim().ToLower().StartsWith("скопје"))
        return true;
    return false;
}

static void NormalizeColumnSize(ISheet sheet)
{
    for (int i = 0; i < 50; i++)
    {
        sheet.AutoSizeColumn(i);
    }
}
