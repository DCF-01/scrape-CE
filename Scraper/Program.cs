﻿using HtmlAgilityPack;
using NPOI;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

const string basePage = @"https://clubeconomy.com.mk";
string filePath = Console.ReadLine().Trim().Replace("/", @"\");
int pageNumber = 1;
bool nextExists = true;

List<List<string>> pageLists = new List<List<string>>();

while (nextExists)
{
    HtmlWeb web = new HtmlWeb();
    var companiesListPage =
            web
            .Load(basePage + "/BusinessAddressBook?pagenumber=" + pageNumber.ToString());

    nextExists = companiesListPage?.DocumentNode.SelectSingleNode(@"//div[@class='no-result']") == null;

    if (nextExists)
    {
        ProcessPage(companiesListPage?.DocumentNode.SelectNodes("//h2/a[@target='_blank']"), ref pageLists);
        Console.WriteLine($"page: {pageNumber}");
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
            var row = worksheet.CreateRow(i + 1);
            for (int j = 0; j < pageLists[i].Count; j++)
            {
                row.CreateCell(j + 1).SetCellValue(pageLists[i][j].ToString());
            }
        }
        NormalizeColumnSize(worksheet);
        // Save to file
        workbook.Write(fs);
    }
}

static void ProcessPage(HtmlNodeCollection? nodes, ref List<List<string>> pageLists)
{
    try
    {
        foreach (var node in nodes)
        {
            var companyTitle = node.InnerText;
            var url = node.Attributes["href"].Value;
            HtmlWeb web = new HtmlWeb();
            var businessPage = web.Load(basePage + url);

            var businessPageNodes = businessPage.DocumentNode.SelectNodes(@"//div[@class='media-body']//a[@target='_blank']");

            var singleCompanyList = new List<string> { companyTitle };
            foreach (var n in businessPageNodes)
            {
                var value = n?.Attributes["href"].Value
                    .Trim()
                    .Replace("mailto:", "")
                    .Replace("tel:", "");

                singleCompanyList.Add(value);
                Console.WriteLine($"{value}");
            }
            pageLists.Add(singleCompanyList);
        }
    }
    catch (Exception e)
    {
        Console.WriteLine(e.Message);
    }
}

static void NormalizeColumnSize(ISheet sheet)
{
    for (int i = 0; i < 50; i++)
    {
        sheet.AutoSizeColumn(i);
    }
}