using IronXL;
using System.Text.RegularExpressions;

public class MyRow
{
    public int Category { get; set; }
    public string Name { get; set; }
    public int Amount { get; set; }
    public decimal Price { get; set; }
    public string CategoryText
    {
        get
        {
            return Category switch
            {
                1 => "Розы(ЕС)",
                2 => "розы куст",
                3 => "гвоздики",
                4 => "хризантемы",
                5 => "зелень",
                6 => "тюльпаны",
                7 => "экзотика",
                8 => "гвоздики",
                _ => "чет не то"
            };
 
        }
    }

    public MyRow(int category, string name, int amount, decimal price)
    {
        Category = category;
        Name = name;
        Amount = amount;
        Price = price;
    }
    public MyRow()
    {

    }

}
public class Program
{
    static void Main()
    {
        List<MyRow> rows = new List<MyRow>();
        List<string> zelen = new List<string>
        {
            "аралия", "берграс", "чико", "петрушка", "феникс", "аспидистра", "писташ", "фисташка","питоспорум","питтоспорум","рускус","эвкалипт","монстера","салал"
        };
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
        // непустые цены вытягиваю из первой
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

        // непустые цены вытягиваю из второй
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
        // непустые количества вытягиваю
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
        var worksheet = workbook.CreateWorkSheet("Итог");
        var worksheet2 = workbook.CreateWorkSheet("Не нашло");

        int p = 1;
        foreach (var pair in flowerAndCount)
        {
            if (pair.Key.Contains("альстр") && !pair.Key.Contains("чарм"))
            {
                var key = flowerAndPrice.Keys.FirstOrDefault(a => a.Contains("альстр") && a.Contains("микс"));
                if (key != null)
                {
                    var value = flowerAndPrice[key];
                    rows.Add(new MyRow(7, pair.Key, pair.Value, value));
                    flowerAndCount.Remove(pair.Key);
                    continue;
                }
            }

            if (pair.Key.Contains("гвозд") && pair.Key.Contains("мини"))
            {
                var key = flowerAndPrice.Keys.FirstOrDefault(a => a.Contains("гвозд") && a.Contains("микс") && a.Contains("мини"));
                if (key != null)
                {
                    var value = flowerAndPrice[key];
                    rows.Add(new MyRow(8, pair.Key, pair.Value, value));
                    flowerAndCount.Remove(pair.Key);
                   
                    continue;
                }
            }
            if (flowerAndPrice.ContainsKey(pair.Key))
            {
                int category = 7;
                if(pair.Key.Contains("роза"))
                { 
                    if(pair.Key.Contains("ветв"))
                        category = 2;
                    else
                        category = 1;
                    
                }
                if (pair.Key.Contains("гвозд"))
                    category = 3;
                if (pair.Key.Contains("хриза"))
                    category = 4;
                if (zelen.Any(pair.Key.Contains))
                    category = 5;
                if(pair.Key.Contains("тюльп"))
                    category = 6;
                rows.Add(new MyRow(category, pair.Key, pair.Value, flowerAndPrice[pair.Key]));
                flowerAndCount.Remove(pair.Key);
               
            }     
        }

       worksheet[$"A{p}"].Value = "РАЗДЕЛ";
       worksheet[$"B{p}"].Value = "КАТЕГОРИЯ ПРЕДЛОЖЕНИЯ";
        worksheet[$"C{p}"].Value = "МЕТКА";
        worksheet[$"D{p}"].Value = "Менеджер по закупу";
        worksheet[$"E{p}"].Value = "Определение Продукта";
        worksheet[$"F{p}"].Value = "Код товара";
        worksheet[$"G{p}"].Value = "Наименование";
        worksheet[$"J{p}"].Value = "ХАРАКТЕРИСТИКА: Страна";
        worksheet[$"K{p}"].Value = "Дата поставки";
        worksheet[$"L{p}"].Value = "Время завершения";
        worksheet[$"M{p}"].Value = "КОЛИЧЕСТВО";
        worksheet[$"N{p}"].Value = "ЛОТ 1";
        worksheet[$"O{p}"].Value = " ЛОТ 2";
        worksheet[$"P{p}"].Value = "ЛОТ 3";
        worksheet[$"Q{p}"].Value = "ЦЕНА 1";
        worksheet[$"R{p}"].Value = "ЦЕНА 2";
        worksheet[$"S{p}"].Value = "ЦЕНА 3";
        worksheet[$"T{p}"].Value = "Закупочная цена";
        p++;
        worksheet[$"A{p}"].Value = "Срезка Импорт";
        p++;
        worksheet[$"B{p}"].Value = "Розы(ЕС)";
        p++;
        foreach(var row in rows.Where(a=>a.Category==1))
        {
            //worksheet[$"D{p}"].Value = "Егорчик";
            worksheet[$"E{p}"].Value = Regex.Match(row.Name, @"^[^0-9]*").Value.Trim();
            worksheet[$"G{p}"].Value = row.Name;
            worksheet[$"K{p}"].Value = DateTime.Now.ToString("MM.dd.yyyy");
            worksheet[$"L{p}"].Value = DateTime.Now.ToString("dd.MM.yyyy tt");
            worksheet[$"M{p}"].Value = row.Amount;
            worksheet[$"N{p}"].Value = 25;
            if (row.Amount >= 100)
            {
                worksheet[$"O{p}"].Value = 100;
                worksheet[$"R{p}"].Value = Convert.ToDecimal((row.Price * 0.97M).ToString("0.##"));
            }
            worksheet[$"Q{p}"].Value = row.Price;
               
            p++;
        }

        worksheet[$"B{p}"].Value = "Розы куст";
        p++;
        foreach (var row in rows.Where(a => a.Category == 2))
        {
            //worksheet[$"D{p}"].Value = "Егорчик";
            worksheet[$"E{p}"].Value = Regex.Match(row.Name, @"^[^0-9]*").Value.Trim();
            worksheet[$"G{p}"].Value = row.Name;
            worksheet[$"K{p}"].Value = DateTime.Now.ToString("MM.dd.yyyy");
            worksheet[$"L{p}"].Value = DateTime.Now.ToString("dd.MM.yyyy tt");
            worksheet[$"M{p}"].Value = row.Amount;
            worksheet[$"N{p}"].Value = 10;
            if (row.Amount >= 50)
            {
                worksheet[$"O{p}"].Value = 50;
                worksheet[$"R{p}"].Value = Convert.ToDecimal((row.Price * 0.95M).ToString("0.##"));
            }
            worksheet[$"Q{p}"].Value = row.Price;

            p++;
        }

        worksheet[$"B{p}"].Value = "Гвоздики";
        p++;
        foreach (var row in rows.Where(a => a.Category == 3 || a.Category == 8))
        {
            worksheet[$"E{p}"].Value = Regex.Match(row.Name, @"^[^0-9]*").Value.Trim();
            worksheet[$"G{p}"].Value = row.Name;
            worksheet[$"K{p}"].Value = DateTime.Now.ToString("MM.dd.yyyy");
            worksheet[$"L{p}"].Value = DateTime.Now.ToString("dd.MM.yyyy tt");
            worksheet[$"M{p}"].Value = row.Amount;
            if (row.Name.Contains("мини"))
            {
                worksheet[$"N{p}"].Value = 10;
                if (row.Amount >= 50)
                {
                    worksheet[$"O{p}"].Value = 50;
                    worksheet[$"R{p}"].Value = Convert.ToDecimal(row.Price.ToString("0.##"));
                }
            }
            else
            {
                worksheet[$"N{p}"].Value = 25;
                if (row.Amount >= 100)
                {
                    worksheet[$"O{p}"].Value = 100;
                    worksheet[$"R{p}"].Value = Convert.ToDecimal(row.Price.ToString("0.##"));
                }
            }

            worksheet[$"Q{p}"].Value = row.Price;

            p++;
        }

        worksheet[$"B{p}"].Value = "Зелень";
        p++;

        foreach (var row in rows.Where(a => a.Category == 5))
        {
            worksheet[$"E{p}"].Value = Regex.Match(row.Name, @"^[^0-9]*").Value.Trim();
            worksheet[$"G{p}"].Value = row.Name;
            worksheet[$"K{p}"].Value = DateTime.Now.ToString("MM.dd.yyyy");
            worksheet[$"L{p}"].Value = DateTime.Now.ToString("dd.MM.yyyy tt");
            worksheet[$"M{p}"].Value = row.Amount;
            if ((row.Name.Contains("питос") || row.Name.Contains("рускус")) && row.Name.Contains("крупн"))
            {
                worksheet[$"N{p}"].Value = 10;
                if (row.Amount >= 50)
                {
                    worksheet[$"O{p}"].Value = 50;
                    worksheet[$"R{p}"].Value = Convert.ToDecimal((row.Price * 0.95M).ToString("0.##"));
                }
            }
            else
            {
                worksheet[$"N{p}"].Value = 1;
                if (row.Amount >= 10)
                {
                    worksheet[$"O{p}"].Value = 10;
                    worksheet[$"R{p}"].Value = Convert.ToDecimal((row.Price * 0.95M).ToString("0.##"));
                }
            }

            worksheet[$"Q{p}"].Value = row.Price;

            p++;
        }

        worksheet[$"B{p}"].Value = "Хризантемы";
        p++;


        foreach (var row in rows.Where(a => a.Category == 4))
        {
            worksheet[$"E{p}"].Value = Regex.Match(row.Name, @"^[^0-9]*").Value.Trim();
            worksheet[$"G{p}"].Value = row.Name;
            worksheet[$"K{p}"].Value = DateTime.Now.ToString("MM.dd.yyyy");
            worksheet[$"L{p}"].Value = DateTime.Now.ToString("dd.MM.yyyy tt");
            worksheet[$"M{p}"].Value = row.Amount;
            if (row.Name.Contains('1'))
            {
                worksheet[$"N{p}"].Value = 10;
                
            }
            else
            {
                worksheet[$"N{p}"].Value = 5;
                if (row.Amount >= 20)
                {
                    worksheet[$"O{p}"].Value = 20;
                    worksheet[$"R{p}"].Value = Convert.ToDecimal(row.Price.ToString("0.##"));
                }
                if (row.Amount >= 40)
                {
                    worksheet[$"O{p}"].Value = 40;
                    worksheet[$"R{p}"].Value = Convert.ToDecimal((row.Price * 0.95M).ToString("0.##"));
                }
            }

            worksheet[$"Q{p}"].Value = row.Price;

            p++;
        }

        worksheet[$"B{p}"].Value = "Тюльпаны";
        p++;

        foreach (var row in rows.Where(a=>a.Category==6))
        {
            worksheet[$"E{p}"].Value = Regex.Match(row.Name, @"^[^0-9]*").Value.Trim();
            worksheet[$"G{p}"].Value = row.Name;
            worksheet[$"K{p}"].Value = DateTime.Now.ToString("MM.dd.yyyy");
            worksheet[$"L{p}"].Value = DateTime.Now.ToString("dd.MM.yyyy tt");
            worksheet[$"M{p}"].Value = row.Amount;
            //worksheet[$"N{p}"].Value = 25;
            worksheet[$"Q{p}"].Value = row.Price;
               
            p++;
        }

        worksheet[$"B{p}"].Value = "Экзотика";
        p++;

        foreach (var row in rows.Where(a => a.Category == 7))
        {
            worksheet[$"E{p}"].Value = Regex.Match(row.Name, @"^[^0-9]*").Value.Trim();
            worksheet[$"G{p}"].Value = row.Name;
            worksheet[$"K{p}"].Value = DateTime.Now.ToString("MM.dd.yyyy");
            worksheet[$"L{p}"].Value = DateTime.Now.ToString("dd.MM.yyyy tt");
            worksheet[$"M{p}"].Value = row.Amount;
            //worksheet[$"N{p}"].Value = 25;
            worksheet[$"Q{p}"].Value = row.Price;

            p++;
        }




        p = 1;
        worksheet2[$"A{p}"].Value = "Ниже не нашло по количеству";
        p++;
        foreach (var pair in flowerAndCount)
        {

            worksheet2[$"A{p}"].Value = pair.Key;
            worksheet2[$"B{p}"].Value = pair.Value;
            flowerAndCount.Remove(pair.Key);
            p++;

        }


        workbook.SaveAs($"{projectDirectory}/excel/result_{DateTimeOffset.UtcNow.ToUnixTimeSeconds()}.xlsx");

    }
}