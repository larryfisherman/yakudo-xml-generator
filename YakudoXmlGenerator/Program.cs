using ClosedXML.Excel;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;

// ── Konfiguracja ─────────────────────────────────────────────────────────────
const string BaseUrl   = "http://localhost:8200";
const string LoginUrl  = "/auth/login";
const string Username  = "admin";
const string Password  = ""; // ← uzupełnij przed uruchomieniem
// ─────────────────────────────────────────────────────────────────────────────

if (args.Length == 0)
{
    Console.WriteLine("Użycie:");
    Console.WriteLine("  dotnet run -- <plik.xlsx>           → generuje pliki XML");
    Console.WriteLine("  dotnet run -- <plik.xlsx> --import  → importuje bezpośrednio do Yakudo");
    return;
}

var xlsxPath   = args[0];
var importMode = args.Contains("--import");

if (!File.Exists(xlsxPath))
{
    Console.WriteLine($"Plik nie istnieje: {xlsxPath}");
    return;
}

Console.OutputEncoding = System.Text.Encoding.UTF8;

if (importMode)
    await RunImport(xlsxPath);
else
    RunXmlExport(xlsxPath);

// ── Tryb HTTP import ──────────────────────────────────────────────────────────

async Task RunImport(string path)
{
    var handler = new HttpClientHandler
    {
        AllowAutoRedirect  = true,
        UseCookies         = true,
        CookieContainer    = new CookieContainer()
    };
    using var http = new HttpClient(handler) { BaseAddress = new Uri(BaseUrl) };
    http.DefaultRequestHeaders.Add("User-Agent", "Mozilla/5.0");

    // Logowanie
    Console.WriteLine("Logowanie...");
    var loginResp0 = await http.GetAsync(LoginUrl);
    if (!loginResp0.IsSuccessStatusCode)
    {
        Console.WriteLine($"Nie znaleziono strony logowania ({LoginUrl}): {loginResp0.StatusCode}");
        Console.WriteLine("Sprawdź stałą LoginUrl w Program.cs (linia 8).");
        return;
    }
    var loginPage  = await loginResp0.Content.ReadAsStringAsync();
    var loginToken = ExtractToken(loginPage);

    if (string.IsNullOrEmpty(loginToken))
    {
        Console.WriteLine($"Nie znaleziono tokenu na stronie logowania. Sprawdź LoginUrl ({LoginUrl}).");
        return;
    }

    var loginResp = await http.PostAsync(LoginUrl, new FormUrlEncodedContent(new Dictionary<string, string>
    {
        ["UsernameOrEmail"]             = Username,
        ["Password"]                    = Password,
        ["__RequestVerificationToken"]  = loginToken
    }));

    var afterLogin = await http.GetAsync("/plus");
    if (afterLogin.RequestMessage?.RequestUri?.AbsolutePath.Contains("login") == true)
    {
        Console.WriteLine("Logowanie nieudane — sprawdź hasło lub adres LoginUrl.");
        return;
    }
    Console.WriteLine("Zalogowano.\n");

    using var workbook = new XLWorkbook(path);
    var sheet = workbook.Worksheet(1);
    // Pobierz token CSRF raz — ważny przez całą sesję
    var addPage = await http.GetStringAsync("/plus/add");
    var token   = ExtractToken(addPage);
    if (string.IsNullOrEmpty(token))
    {
        Console.WriteLine("Nie znaleziono tokenu CSRF na stronie /plus/add.");
        return;
    }

    int imported = 0, skipped = 0, failed = 0;

    foreach (var row in sheet.RowsUsed().Skip(1))
    {
        if (row.Cell(11).IsEmpty()) { skipped++; continue; }

        var plu      = CellToString(row.Cell(11));
        var nazwa    = CellToString(row.Cell(12));
        var ean      = ComputeEan(CellToString(row.Cell(2)), plu);
        var price    = row.Cell(5).IsEmpty() ? 0.0 : row.Cell(5).GetValue<double>();
        var priceStr = price.ToString("F2", System.Globalization.CultureInfo.InvariantCulture);

        var form     = BuildForm(nazwa, plu, ean, priceStr, token);
        var resp     = await http.PostAsync("/plus/add", form);
        var finalUrl = resp.RequestMessage?.RequestUri?.AbsolutePath ?? "";
        bool ok      = finalUrl == "/plus" || finalUrl.StartsWith("/plus?");

        if (ok)
        {
            Console.WriteLine($"  [OK]  PLU {plu,-5}  {nazwa}");
            imported++;
        }
        else
        {
            Console.WriteLine($"  [ERR] PLU {plu,-5}  {nazwa}  (status: {resp.StatusCode}, url: {finalUrl})");
            if (failed == 0)
            {
                var body = await resp.Content.ReadAsStringAsync();
                File.WriteAllText("error_response.html", body, System.Text.Encoding.UTF8);
                Console.WriteLine("       Zapisano error_response.html - otworz w przegladarce zeby zobaczyc blad");
            }
            failed++;
        }

        await Task.Delay(300);
    }

    Console.WriteLine($"\nGotowe: {imported} zaimportowanych, {failed} błędów, {skipped} pominięto.");
}

MultipartFormDataContent BuildForm(string nazwa, string plu, string ean, string price, string token)
{
    var f = new MultipartFormDataContent();
    void A(string name, string val) => f.Add(new StringContent(val), name);

    A("Item.Id",                        "0");
    A("Item.Name",                      nazwa);
    A("PluDepartmentId",                "1");
    A("Item.GroupId",                   "1");
    A("Item.PluNumber",                 plu);
    A("Item.WeightingMode",             "Weighted");
    A("Item.UnitPrice",                 price);
    A("Item.Tare",                      "0");
    A("Item.BarcodeDefinitionId",       "");
    A("Item.BarcodeFormat",             "f1F2_CCCC_XXXXXX_CD");
    A("Item.Ean",                       ean);
    A("Item.BarcodeFormatRightSideData","price");
    A("Item.PluDiscountType",           "NoDiscount");
    A("Item.PluDiscountDays",           "None");
    A("Item.PluDiscountForFirstLimit",  "");
    A("Item.PrintedName.Text",          "");
    A("Item.PrintedName.Fonts",         "");
    A("Item.Ingredients.Text",          "");
    for (int i = 0; i < 15; i++) A("Item.Ingredients.Fonts", "");
    A("Item.AdditionalText.Text",       "");
    for (int i = 0; i < 5;  i++) A("Item.AdditionalText.Fonts", "");
    A("Item.StorageTemperatureMin",     "");
    A("Item.StorageTemperatureMax",     "");
    A("Item.UseByDate",                 "");
    A("Item.SellByDate",                "");
    A("Item.SellByDateSource",          "1");
    A("Item.PackByDate",                "0");
    A("Item.SellByTimeEnabled",         "false");
    A("Item.PackByTimeEnabled",         "true");
    A("Item.PackByTime",                "");
    A("Item.ProductionPlaceId",         "");
    A("Item.MinWeight",                 "");
    A("Item.MaxWeight",                 "");
    A("Item.DisplayImageNumber",        "");
    A("DisplayImage.Name",              "");
    A("DisplayImage.Path",              "");
    A("DisplayImage.Delete",            "false");
    A("DisplayImage.ContentType",       "");
    A("Item.KeyboardName",              "");
    A("Item.ContainerNumber",           "");

    // LabelFormats
    A("Item.LabelFormats[0].LabelType", "Normal");
    A("Item.LabelFormats[0].ProfileId", "1");
    A("Item.LabelFormats[0].Ordinal",   "1");
    A("Item.LabelFormats[0].Value",     "1");
    A("Item.LabelFormats[1].LabelType", "Normal");
    A("Item.LabelFormats[1].ProfileId", "1");
    A("Item.LabelFormats[1].Ordinal",   "2");
    A("Item.LabelFormats[1].Value",     "-1");

    // Texts (5 wpisów)
    for (int i = 0; i < 5; i++)
    {
        A($"Item.Texts[{i}].ProfileId", "1");
        A($"Item.Texts[{i}].Ordinal",   (i + 1).ToString());
        A($"Item.Texts[{i}].Value",     "-1");
    }

    // Images (10 wpisów)
    for (int i = 0; i < 10; i++)
    {
        A($"Item.Images[{i}].ProfileId", "1");
        A($"Item.Images[{i}].Ordinal",   (i + 1).ToString());
        A($"Item.Images[{i}].Value",     "-1");
    }

    A("AllergenFormattingStyle",            "none");
    A("__RequestVerificationToken",         token);
    A("Item.PreventDelete",                 "false");
    A("Item.UsePluAddionalBarcodeData",     "false");
    A("Item.UseNutritionFacts",             "false");
    A("Item.UseTraceability",               "false");

    return f;
}

// ── Tryb XML export (bez zmian) ───────────────────────────────────────────────

void RunXmlExport(string path)
{
    var outputDir = "xml_output";
    Directory.CreateDirectory(outputDir);
    var timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
    int generated = 0, skipped = 0;

    using var workbook = new XLWorkbook(path);
    var sheet = workbook.Worksheet(1);

    foreach (var row in sheet.RowsUsed().Skip(1))
    {
        if (row.Cell(11).IsEmpty()) { skipped++; continue; }

        var plu   = CellToString(row.Cell(11));
        var nazwa = CellToString(row.Cell(12));
        var ean   = ComputeEan(CellToString(row.Cell(2)), plu);
        var price = row.Cell(5).IsEmpty() ? 0.0 : row.Cell(5).GetValue<double>();

        var xml      = BuildXml(nazwa, plu, ean, price);
        var fileName = $"Plu_{plu}_{ean}_{timestamp}.xml";
        File.WriteAllText(Path.Combine(outputDir, fileName), xml, new UTF8Encoding(false));
        Console.WriteLine($"  {fileName}");
        generated++;
    }

    Console.WriteLine($"\nGotowe: {generated} plików w katalogu '{outputDir}'");
    if (skipped > 0) Console.WriteLine($"Pominięto {skipped} wierszy (brak PLU).");
}

// ── Helpers ───────────────────────────────────────────────────────────────────

static string ExtractToken(string html)
{
    var m = Regex.Match(html, @"name=""__RequestVerificationToken""[^>]*value=""([^""]+)""");
    if (!m.Success)
        m = Regex.Match(html, @"value=""([^""]+)""[^>]*name=""__RequestVerificationToken""");
    return m.Success ? m.Groups[1].Value : "";
}

static string CellToString(IXLCell cell)
{
    if (cell.IsEmpty()) return "";
    return cell.Value.IsNumber
        ? ((long)cell.Value.GetNumber()).ToString()
        : cell.GetValue<string>().Trim();
}

static string ComputeEan(string raw, string plu)
{
    if (!string.IsNullOrEmpty(raw) && raw.Length == 13) return raw;
    return ("290" + plu).PadRight(13, '0');
}

static string XmlEscape(string s) => s
    .Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;");

static string BuildXml(string name, string plu, string ean, double price)
{
    var p = price.ToString("F3", System.Globalization.CultureInfo.InvariantCulture);
    return
$"""
<?xml version="1.0"?>
<Plu xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema">
  <Name>{XmlEscape(name)}</Name>
  <PluNumber>{plu}</PluNumber>
  <Ean>{ean}</Ean>
  <GroupId>1</GroupId>
  <InternalNumber xsi:nil="true" />
  <WeightingMode>Weighted</WeightingMode>
  <UnitPrice>{p}</UnitPrice>
  <UnitPriceBase>0</UnitPriceBase>
  <Tare>0</Tare>
  <Quantity>1</Quantity>
  <QuantitySymbol>Pcs</QuantitySymbol>
  <UseUseByDate>false</UseUseByDate>
  <UseByDate xsi:nil="true" />
  <SellByDate xsi:nil="true" />
  <SellByTimeUseRTC>false</SellByTimeUseRTC>
  <SellByTime xsi:nil="true" />
  <PackByDate>0</PackByDate>
  <PackByTimeUseRTC>false</PackByTimeUseRTC>
  <PackByTime xsi:nil="true" />
  <StorageTemperatureMin xsi:nil="true" />
  <StorageTemperatureMax xsi:nil="true" />
  <PrintedName><Line><Number>1</Number><Text /><Font /></Line></PrintedName>
  <AdditionalText><Line><Number>1</Number><Text /><Font /></Line></AdditionalText>
  <Ingredients><Line><Number>1</Number><Text /><Font /></Line></Ingredients>
  <UseNutritionFacts>false</UseNutritionFacts>
  <BarcodeFormat>F1F2_CCCC_XXXXXX_CD</BarcodeFormat>
  <BarcodeFormatRightSideData>Price</BarcodeFormatRightSideData>
  <DisplayImageNumber xsi:nil="true" />
  <AutoRecalculateNutritions>true</AutoRecalculateNutritions>
  <UseTraceability>false</UseTraceability>
  <NutritionFacts />
  <LabelFormats>
    <PluRelatedDataDto><ProfileName>Domyślny</ProfileName><Value>1</Value><Ordinal>1</Ordinal></PluRelatedDataDto>
    <PluRelatedDataDto><ProfileName>Domyślny</ProfileName><Value>-1</Value><Ordinal>2</Ordinal></PluRelatedDataDto>
  </LabelFormats>
  <Images>
    <PluRelatedDataDto><ProfileName>Domyślny</ProfileName><Value>-1</Value><Ordinal>1</Ordinal></PluRelatedDataDto>
    <PluRelatedDataDto><ProfileName>Domyślny</ProfileName><Value>-1</Value><Ordinal>2</Ordinal></PluRelatedDataDto>
    <PluRelatedDataDto><ProfileName>Domyślny</ProfileName><Value>-1</Value><Ordinal>3</Ordinal></PluRelatedDataDto>
    <PluRelatedDataDto><ProfileName>Domyślny</ProfileName><Value>-1</Value><Ordinal>4</Ordinal></PluRelatedDataDto>
    <PluRelatedDataDto><ProfileName>Domyślny</ProfileName><Value>-1</Value><Ordinal>5</Ordinal></PluRelatedDataDto>
    <PluRelatedDataDto><ProfileName>Domyślny</ProfileName><Value>-1</Value><Ordinal>6</Ordinal></PluRelatedDataDto>
    <PluRelatedDataDto><ProfileName>Domyślny</ProfileName><Value>-1</Value><Ordinal>7</Ordinal></PluRelatedDataDto>
    <PluRelatedDataDto><ProfileName>Domyślny</ProfileName><Value>-1</Value><Ordinal>8</Ordinal></PluRelatedDataDto>
    <PluRelatedDataDto><ProfileName>Domyślny</ProfileName><Value>-1</Value><Ordinal>9</Ordinal></PluRelatedDataDto>
    <PluRelatedDataDto><ProfileName>Domyślny</ProfileName><Value>-1</Value><Ordinal>10</Ordinal></PluRelatedDataDto>
  </Images>
  <Texts>
    <PluRelatedDataDto><ProfileName>Domyślny</ProfileName><Value>-1</Value><Ordinal>1</Ordinal></PluRelatedDataDto>
    <PluRelatedDataDto><ProfileName>Domyślny</ProfileName><Value>-1</Value><Ordinal>2</Ordinal></PluRelatedDataDto>
    <PluRelatedDataDto><ProfileName>Domyślny</ProfileName><Value>-1</Value><Ordinal>3</Ordinal></PluRelatedDataDto>
    <PluRelatedDataDto><ProfileName>Domyślny</ProfileName><Value>-1</Value><Ordinal>4</Ordinal></PluRelatedDataDto>
    <PluRelatedDataDto><ProfileName>Domyślny</ProfileName><Value>-1</Value><Ordinal>5</Ordinal></PluRelatedDataDto>
  </Texts>
</Plu>
""";
}
