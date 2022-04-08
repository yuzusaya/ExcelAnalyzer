using OfficeOpenXml;

Console.WriteLine("Please input the file path:");
var path = Console.ReadLine();
path = path.Replace("\"", "");
ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

using (var package = new ExcelPackage(new FileInfo(path)))
{
    var worksheet = package.Workbook.Worksheets.First();
    var rowCount = worksheet.Dimension.Rows;
    var columnCount = worksheet.Dimension.Columns;

    Console.WriteLine("Please input the columns that you want to search (separated by comma E.g.:(A,C,D)");
    var inputColumns = Console.ReadLine().Split(",");
    List<int> columns = new List<int>();
    foreach (var inputColumn in inputColumns)
    {
        columns.Add(char.ToUpper(inputColumn.Trim().First()) - 64);
    }

    Console.WriteLine($"Please input the keyword(s) for column {inputColumns.First()} (separated by comma if multiple)");
    var keyword = Console.ReadLine();
    var keywords = keyword.Split(',');

    var valuesNeedToMatch = new List<string>();
    for (int i = 1; i < columns.Count; i++)
    {
        Console.WriteLine($"Please input the value that must match for column {inputColumns[i]}");
        valuesNeedToMatch.Add(Console.ReadLine());
    }

    var count = 0;
    for (int row = 1; row <= rowCount; row++)
    {
        var value = (worksheet.Cells[row, columns.First()].Value?.ToString())??string.Empty;
        var values = value.Split(',');
        var match = true;
        for (int i = 0; i < keywords.Length; i++)
        {
            if (values.Any(x => x.Trim().Equals(keywords[i].Trim())))
            {
                for (int j = 1; j < columns.Count; j++)
                {
                    var valueMustMatch = worksheet.Cells[row, columns[j]].Value?.ToString().Trim().Replace("'","");
                    if (!valueMustMatch.Equals(valuesNeedToMatch[j - 1]))
                    {
                        match = false;
                        break;
                    }
                }

                if (match)
                {
                    count++;
                    break;
                }
            }
        }
    }

    Console.WriteLine($"Match count: {count}");






    //List<int> counts = new List<int>();
    //foreach (var key in keywords)
    //{
    //    counts.Add(0);
    //}
    //Console.WriteLine("Please specify the column to search the keyword");
    //var column = Console.ReadKey().KeyChar;
    //var col = char.ToUpper(column) - 64;
    //for (int row = 1; row <= rowCount; row++)
    //{
    //    var value = worksheet.Cells[row, col].Value?.ToString();
    //    var values = value.Split(',');
    //    for (int i = 0; i < keywords.Length; i++)
    //    {
    //        if (values.Any(x => x.Trim().Equals(keywords[i].Trim())))
    //        {
    //            counts[i]++;
    //        }
    //    }
    //}

    //for (int i = 0; i < keywords.Length; i++)
    //{
    //    Console.WriteLine($"key:{keywords[i]}, count:{counts[i]}");

    //}

    Console.ReadLine();
}