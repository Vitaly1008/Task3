// See https://aka.ms/new-console-template for more information
using ClosedXML.Excel;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Globalization;
using System.IO.Enumeration;
using System.IO.Pipes;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices.Marshalling;
using Task3;


Console.WriteLine("Приложение для работы с excel таблицей.");

// списки с данными, загруженными из таблицы эксель
List<Product> products = new List<Product>();
List<Сlient> сlients = new List<Сlient>();
List<Request> requests = new List<Request>();
string fileName = "Приложение2.xlsx", fullFileName = "";

fullFileName = ShowDialogToOpenFile(fileName, products, сlients, requests);
ShowConsoleDialog(products, сlients, requests, fullFileName);

static void ShowConsoleDialog(List<Product> products, List<Сlient> сlients, List<Request> requests, string fileName = "")
{
    Console.WriteLine("Введите 1 для вывода информации о клиенте по наименованию товара (задача№2)");
    Console.WriteLine("Введите 2 изменения ФИО для организации (задача№3)");
    Console.WriteLine("Введите 3 для вывода информации о золотом клиенте (задача №4)");
    Console.WriteLine("Введите exit для выхода из программы");
    string choise = Console.ReadLine();
    switch (choise)
    {
        case "1":
            Task2(products, сlients, requests);
            ShowConsoleDialog(products, сlients, requests);
            break;
        case "2":
            Task3(products, сlients, requests, fileName);
            ShowConsoleDialog(products, сlients, requests);
            break;
        case "3":
            Task4(products, сlients, requests);
            ShowConsoleDialog(products, сlients, requests);
            break;
        case "exit":
            Exit();
            break;
        default:
            ShowConsoleDialog(products, сlients, requests);
            break;
    }
}
static void Task2(List<Product> products, List<Сlient> сlients, List<Request> requests)
{
    Console.WriteLine("Введите наименование товара");
    string searchProduct = Console.ReadLine();

    Product serchingProduct = products.Find(x => x.Name == searchProduct);
    if (serchingProduct != null)
    {
        List<Request> serchingRequest = requests.FindAll(x => x.ProductCode == serchingProduct.ProductCode);
        if (serchingRequest != null)
        {
            foreach (Request rqst in serchingRequest)
            {
                Сlient serchingclient = сlients.Find(x => x.ClientCode == rqst.ClientCode);
                if (serchingclient != null)
                    Console.WriteLine($"Товар {serchingProduct.Name} был заказан клиентом с кодом {serchingclient.ClientCode},\n" +
                        $"названием организации {serchingclient.Organization}, по адресу {serchingclient.Address} \n" +
                        $"и с ФИО контакта {serchingclient.Contact}.\n Цена товара {serchingProduct.UnitPrice}, в количестве {rqst.Quantity}, {rqst.Date} числа \n");
                else Console.WriteLine("Запрос на покупку данного товара не найден.\n");
            }
        } else { Console.WriteLine("Запрос на покупку данного товара не найден.\n"); }
    } else { Console.WriteLine("Запрос на покупку данного товара не найден.\n"); }
}
static void Task3(List<Product> products, List<Сlient> сlients, List<Request> requests, string fileName)
{
    Console.WriteLine("Выберите сотрудника по коду сотрудника");
    foreach (var cl in сlients)
        Console.WriteLine($"Код клиента: {cl.ClientCode}, ФИО сотрудника : {cl.Contact}, организация: {cl.Organization} \n");
    string answere = Console.ReadLine();
    int answereClientCode = 0;
    if (int.TryParse(answere, out answereClientCode))
    {
        Сlient cl = сlients.Find(c => c.ClientCode == answereClientCode);
        if (cl != null)
        {
            Console.WriteLine("Введите новую фамилию, имя и отчество");
            answere = Console.ReadLine();
            for (uint i = 0; i < сlients.Count; i++)
            {
                if (сlients[(int)i].ClientCode == answereClientCode)
                { 
                    UpdateExcelUsingOpenXMLSDK(fileName, answere, i+2, "D");
                    Console.WriteLine($"ФИО контрагента изменилось с {сlients[(int)i].Contact} на {answere}");
                }
            }

        } else { Console.WriteLine($"Клиент с кодом {answereClientCode} не найден"); }
    }
    else { Console.WriteLine("Неверный код клиента"); }
}
static void Task4(List<Product> products, List<Сlient> сlients, List<Request> requests) 
{
    Console.WriteLine("Поиск золотого клиента по количеству заказов за указанный период времени");
    Console.WriteLine("Введите год поиска");
    string answere = "";
    answere = Console.ReadLine();
    int searchingYear = 0, searchMounth = 0;
    if (int.TryParse(answere, out searchingYear))
    {
        if (searchingYear <= DateTime.Today.Year && searchingYear >= DateTime.UnixEpoch.Year)
        {
            Console.WriteLine("Введите номер месяца (0 - для поиска по всему году)");
            answere = Console.ReadLine();
            if (int.TryParse(answere, out searchMounth))
            {
                if (searchMounth >= 0 && searchMounth <= 12)
                {
                    DateTime startDate = searchMounth == 0 ? new DateTime(searchingYear, 1, 1) : new DateTime(searchingYear, searchMounth, 1);
                    DateTime endDate = searchMounth == 0 ? new DateTime(searchingYear, 12, 31) : new DateTime(searchingYear, searchMounth+1, 1).AddDays(-1);
                    Dictionary<Сlient, int> quantityPerClient = new Dictionary<Сlient, int>();
                    foreach (Сlient cl in сlients)
                    {
                        quantityPerClient.Add(cl, 0);
                        foreach (Request rqst in requests)
                        {
                            if (cl.ClientCode == rqst.ClientCode && rqst.Date <= endDate && rqst.Date >= startDate)
                                quantityPerClient[cl] += rqst.Quantity;
                        }
                    }
                    // сортировка словаря, надо взять последний
                    quantityPerClient = quantityPerClient.OrderBy(el => el.Value).ToDictionary(el => el.Key, el => el.Value);
                    if (quantityPerClient.Last().Value > 0)
                    {
                        Console.WriteLine($"За текущий период с {startDate} по {endDate} больше всего заказов {quantityPerClient.Last().Value} " +
                        $" клиента с кодом {quantityPerClient.Last().Key.ClientCode} и названием организации {quantityPerClient.Last().Key.Organization}. ");
                    } else { Console.WriteLine("За указанный период не было заявок."); }
                }
                else { Console.WriteLine("Неправильный месяц"); }
            } else { Console.WriteLine("Неправильный месяц"); }
        } else { Console.WriteLine("Неправильный год"); }
    } else { Console.WriteLine("Неправильный год"); }
}
static void Exit()
{
    Console.WriteLine("Завершение работы программы");
    Console.ReadKey();
    Environment.Exit(0);
}
#region функция для открытия файла и извлечением информации
static string ShowDialogToOpenFile(string fileName, List<Product> products, List<Сlient> сlients, List<Request> requests)
{
    Console.WriteLine("Введите путь до файла. Файл должен называться 'Приложение2.xlsx'");
    string fullFilePath = "";
    string pathToFile = Console.ReadLine();
    //pathToFile = @"C:\Users\Админ\Desktop\test\";
    if (Directory.Exists(pathToFile))
    {
        try
        {
            fullFilePath = pathToFile.Substring(pathToFile.Length - 1).Equals(@"\") ? pathToFile + fileName : pathToFile + @"\" + fileName;
            ExecuteDataFromFile(fullFilePath, products, сlients, requests);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Не удалось открыть файл по пути {pathToFile} с названием {fileName}");
            throw;
        }
    }
    else
    {
        Console.WriteLine($"Указан неверный путь к фалу: {pathToFile}");
        Exit();
    }
    return fullFilePath;
}
#endregion

#region функция считывания с эксель документа с определенной страницы, возвращает список со значениями в каждой ячейке
static List<string> GetCellsValue(string fileName, string sheetName)
{
    string? value = null;
    List<string> result = new List<string>();
    // Open the spreadsheet document for read-only access.
    using (SpreadsheetDocument document = SpreadsheetDocument.Open(fileName, false))
    {
        // Retrieve a reference to the workbook part.
        WorkbookPart? wbPart = document.WorkbookPart;
        // Find the sheet with the supplied name, and then use that 
        // Sheet object to retrieve a reference to the first worksheet.
        Sheet? theSheet = wbPart?.Workbook.Descendants<Sheet>().Where(s => s.Name == sheetName).FirstOrDefault();

        // Throw an exception if there is no sheet.
        if (theSheet is null || theSheet.Id is null)
        {
            throw new ArgumentException("sheetName");
        }
        // Retrieve a reference to the worksheet part.
        WorksheetPart wsPart = (WorksheetPart)wbPart!.GetPartById(theSheet.Id!);
        // Use its Worksheet property to get a reference to the cell 
        // whose address matches the address you supplied.
        foreach (Cell theCell in wsPart.Worksheet?.Descendants<Cell>())
        {
            if (theCell.InnerText.Length > 0)
            {
                value = theCell.InnerText;
                // If the cell represents an integer number, you are done. 
                // For dates, this code returns the serialized value that 
                // represents the date. The code handles strings and 
                // Booleans individually. For shared strings, the code 
                // looks up the corresponding value in the shared string 
                // table. For Booleans, the code converts the value into 
                // the words TRUE or FALSE.
                if (theCell.DataType is not null)
                {
                    if (theCell.DataType.Value == CellValues.SharedString)
                    {
                        // For shared strings, look up the value in the
                        // shared strings table.
                        var stringTable = wbPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
                        // If the shared string table is missing, something 
                        // is wrong. Return the index that is in
                        // the cell. Otherwise, look up the correct text in 
                        // the table.
                        if (stringTable is not null)
                        {
                            value = stringTable.SharedStringTable.ElementAt(int.Parse(value)).InnerText;
                        }
                    }
                    else if (theCell.DataType.Value == CellValues.Boolean)
                    {
                        switch (value)
                        {
                            case "0":
                                value = "FALSE";
                                break;
                            default:
                                value = "TRUE";
                                break;
                        }
                    }
                }
                result.Add(value);
            }
        }
    }

    return result;
}
#endregion

#region функция извлечение данных с файла
static void ExecuteDataFromFile(string? pathToFile, List<Product> products, List<Сlient> сlients, List<Request> requests)
{
    //заполняем 
    List<string> dataFromFile = GetCellsValue(pathToFile, "Товары");
    for (int i = 1; i < dataFromFile.Count / 4; i++)
    {
        MeasureEnum measure = MeasureEnum.none;
        switch (dataFromFile[4 * i + 2])
        {
            case "Литр":
                measure = MeasureEnum.liter; break;
            case "Килограмм":
                measure = MeasureEnum.kilogram; break;
            case "Штука":
                measure = MeasureEnum.piece; break;
            default:
                measure = MeasureEnum.none; break;
        }
        products.Add(new Product(Convert.ToInt32(dataFromFile[4 * i]),
                                    dataFromFile[4 * i + 1].ToString(),
                                    measure,
                                    float.Parse(dataFromFile[4 * i + 3])));
    }

    dataFromFile = GetCellsValue(pathToFile, "Клиенты");
    for (int i = 1; i < dataFromFile.Count / 4; i++)
    {
        сlients.Add(new Сlient(Convert.ToInt32(dataFromFile[4 * i]),
                                dataFromFile[4 * i + 1],
                                dataFromFile[4 * i + 2],
                                dataFromFile[4 * i + 3]));
    }

    dataFromFile = GetCellsValue(pathToFile, "Заявки");
    for (int i = 1; i < dataFromFile.Count / 6; i++)
    {
        var excelZeroDateTime = new DateTime(1899, 12, 30);
        var dt = excelZeroDateTime.AddDays(Convert.ToUInt32(dataFromFile[6 * i + +5]));
        requests.Add(new Request(Convert.ToInt32(dataFromFile[6 * i]),
                                    Convert.ToInt32(dataFromFile[6 * i + 1]),
                                    Convert.ToInt32(dataFromFile[6 * i + 2]),
                                    Convert.ToInt32(dataFromFile[6 * i + 3]),
                                    Convert.ToInt32(dataFromFile[6 * i + 4]),
                                    dt));
    }
}
#endregion

#region Вставка текста
static void UpdateExcelUsingOpenXMLSDK(string fileName, string newName, uint i, string j)
{
    // Open the document for editing.
    using (SpreadsheetDocument spreadSheet = SpreadsheetDocument.Open(fileName, true))
    {
        // Access the main Workbook part, which contains all references.
        WorkbookPart workbookPart = spreadSheet.WorkbookPart;
        // get sheet by name
        Sheet sheet = workbookPart.Workbook.Descendants<Sheet>().Where(s => s.Name == "Клиенты").FirstOrDefault();
        // get worksheetpart by sheet id
        WorksheetPart worksheetPart = workbookPart.GetPartById(sheet.Id.Value) as WorksheetPart;
        // The SheetData object will contain all the data.
        SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
        Cell cell = GetCell(worksheetPart.Worksheet, j, i);
        cell.CellValue = new CellValue(newName);
        cell.DataType = new EnumValue<CellValues>(CellValues.String);
        // Save the worksheet.
        worksheetPart.Worksheet.Save();
        // for recacluation of formula
        spreadSheet.WorkbookPart.Workbook.CalculationProperties.ForceFullCalculation = true;
        spreadSheet.WorkbookPart.Workbook.CalculationProperties.FullCalculationOnLoad = true;
    }
}

static Cell GetCell(Worksheet worksheet,
string columnName, uint rowIndex)
{
    Row row = GetRow(worksheet, rowIndex);
    if (row == null) return null;
    var FirstRow = row.Elements<Cell>().Where(c => string.Compare
    (c.CellReference.Value, columnName +
    rowIndex, true) == 0).FirstOrDefault();
    if (FirstRow == null) return null;
    return FirstRow;
}

static Row GetRow(Worksheet worksheet, uint rowIndex)
{
    Row row = worksheet.GetFirstChild<SheetData>().
    Elements<Row>().FirstOrDefault(r => r.RowIndex == rowIndex);
    if (row == null)
    {
        throw new ArgumentException(String.Format("No row with index {0} found in spreadsheet", rowIndex));
    }
    return row;
}
#endregion

