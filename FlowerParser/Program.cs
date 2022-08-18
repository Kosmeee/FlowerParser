using IronXL;

public class Program
{
    static void Main()
    {
        string workingDirectory = Environment.CurrentDirectory;
        string projectDirectory = Directory.GetParent(workingDirectory).Parent.Parent.FullName;
        string flowerAndPricePath = $"{projectDirectory}/excel/1.xlsx";
        string secondFlowerAndPricePath = $"{projectDirectory}/excel/2.xlsx";
        string flowerAndCountPath = $"{projectDirectory}/excel/3.xlsx";

        WorkBook flowerAndPriceTable = WorkBook.Load(flowerAndPricePath);
        WorkBook secondFlowerAndPriceTable = WorkBook.Load(secondFlowerAndPricePath);
        WorkBook flowerAndCountTable = WorkBook.Load(flowerAndCountPath);

        var flowerAndPriceSheet = flowerAndPriceTable.WorkSheets.First();
        var secondFlowerAndPriceSheet = secondFlowerAndPriceTable.WorkSheets.First();
        var flowerAndCountSheet = flowerAndCountTable.WorkSheets.First();

        int j = 0;
        Dictionary<string, decimal> flowerAndPrice = new();
        for(int i= 1; ;i++)
        {
        
            if (flowerAndPrice.ContainsKey(flowerAndPriceSheet[$"B{i}"].StringValue.Trim().ToLower()) | flowerAndPriceSheet[$"B{i}"].IsEmpty)
            {
                j++;
                if (j > 10) break;
                continue;
                
            }
            flowerAndPrice.Add(flowerAndPriceSheet[$"B{i}"].StringValue.Trim().ToLower(), flowerAndPriceSheet[$"D{i}"].DecimalValue);
            j = 0;
        }
        j = 0;
        for (int i = 1; ; i++)
        {

            if (flowerAndPrice.ContainsKey(secondFlowerAndPriceSheet[$"B{i}"].StringValue.Trim().ToLower()) | secondFlowerAndPriceSheet[$"C{i}"].IsEmpty)
            {
                j++;
                if(j> 10) break;
                continue;
            }
            flowerAndPrice.Add(secondFlowerAndPriceSheet[$"B{i}"].StringValue.Trim().ToLower(), secondFlowerAndPriceSheet[$"C{i}"].DecimalValue);
             j = 0;
        }

        Dictionary<string, int> flowerAndCount = new();
        j = 0;
        for (int i = 1;; i++)
        {
            if (flowerAndCount.ContainsKey(flowerAndCountSheet[$"A{i}"].StringValue.Trim().ToLower()) | flowerAndCountSheet[$"E{i}"].IsEmpty | flowerAndCountSheet[$"E{i}"].IntValue == 0 | flowerAndCountSheet[$"A{i}"].StringValue.Trim().ToLower().Contains("итого"))
            {
                j++;
                if (j > 500) break;
                continue;
            }
            flowerAndCount.Add(flowerAndCountSheet[$"A{i}"].StringValue.Trim().ToLower(), flowerAndCountSheet[$"E{i}"].IntValue);
            j = 0;
            

        }

        var workbook = new WorkBook(ExcelFileFormat.XLSX);
        var worksheet = workbook.CreateWorkSheet("Flowers");

        int p = 1;
        foreach (var pair in flowerAndPrice)
        {
            if(flowerAndCount.ContainsKey(pair.Key))
            {
                worksheet[$"A{p}"].Value = pair.Key;
                worksheet[$"B{p}"].Value = flowerAndCount[pair.Key];
                worksheet[$"C{p}"].Value = pair.Value;
                flowerAndPrice.Remove(pair.Key);
                flowerAndCount.Remove(pair.Key);
                p++;
            }     
        }
        worksheet[$"A{p}"].Value = "Ниже не нашло по ценам";
        p++;
        foreach (var pair in flowerAndPrice)
        {
           
                worksheet[$"A{p}"].Value = pair.Key;
                worksheet[$"B{p}"].Value = pair.Value;
                flowerAndPrice.Remove(pair.Key);
                p++;
            
        }

        worksheet[$"A{p}"].Value = "Ниже не нашло по количеству";
        p++;
        foreach (var pair in flowerAndCount)
        {

            worksheet[$"A{p}"].Value = pair.Key;
            worksheet[$"B{p}"].Value = pair.Value;
            flowerAndCount.Remove(pair.Key);
            p++;

        }


        workbook.SaveAs($"{projectDirectory}/excel/result_{DateTimeOffset.UtcNow.ToUnixTimeSeconds()}.xlsx");

    }
}