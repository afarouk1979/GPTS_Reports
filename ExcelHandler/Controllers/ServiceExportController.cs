using ClosedXML.Excel;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using System.Data;
using System.IO;
using System.Reflection;
using System.Data.Common;
using System.Linq;

[Route("api/[controller]")]
[ApiController]
public class ServiceExportController : ControllerBase
{
    private readonly AppDbContext _context;

    public ServiceExportController(AppDbContext context)
    {
        _context = context;
    }

    #region Tafkeet (Number to Words) Helper
    public static class TafkeetHelper
    {
        public static string ToWords(decimal number)
        {
            if (number == 0)
                return "صفر";

            if (number < 0)
                return "سالب " + ToWords(Math.Abs(number));

            string words = "";

            long intPart = (long)number;
            decimal fractionPart = number - intPart;

            words = ConvertIntegerToWords(intPart);

            if (fractionPart > 0)
            {
                long fractionalValue = (long)Math.Round(fractionPart * 100);
                if (fractionalValue > 0)
                {
                    if (!string.IsNullOrEmpty(words))
                        words += " و ";

                    words += ConvertIntegerToWords(fractionalValue) + " قرشاً";
                }
            }

            return words.Trim();
        }

        private static string ConvertIntegerToWords(long number)
        {
            if (number == 0) return "";

            string[] ones = { "", "واحد", "اثنان", "ثلاثة", "أربعة", "خمسة", "ستة", "سبعة", "ثمانية", "تسعة", "عشرة", "أحد عشر", "اثنا عشر", "ثلاثة عشر", "أربعة عشر", "خمسة عشر", "ستة عشر", "سبعة عشر", "ثمانية عشر", "تسعة عشر" };
            string[] tens = { "", "", "عشرون", "ثلاثون", "أربعون", "خمسون", "ستون", "سبعون", "ثمانون", "تسعون" };
            string[] hundreds = { "", "مائة", "مائتان", "ثلاثمائة", "أربعمائة", "خمسمائة", "ستمائة", "سبعمائة", "ثمانمائة", "تسعمائة" };

            string words = "";
            bool needsWaw = false;

            if (number >= 1000000)
            {
                long millions = number / 1000000;
                if (millions == 1) words += "مليون";
                else if (millions == 2) words += "مليونان";
                else if (millions >= 3 && millions <= 10) words += ConvertIntegerToWords(millions) + " ملايين";
                else words += ConvertIntegerToWords(millions) + " مليون";
                number %= 1000000;
                needsWaw = true;
            }

            if (number >= 1000)
            {
                if (needsWaw && number > 0) words += " و ";
                long thousands = number / 1000;
                if (thousands == 1) words += "ألف";
                else if (thousands == 2) words += "ألفان";
                else if (thousands >= 3 && thousands <= 10) words += ConvertIntegerToWords(thousands) + " آلاف";
                else words += ConvertIntegerToWords(thousands) + " ألف";
                number %= 1000;
                needsWaw = true;
            }

            if (number >= 100)
            {
                if (needsWaw && number > 0) words += " و ";
                words += hundreds[number / 100];
                number %= 100;
                needsWaw = true;
            }

            if (number > 0)
            {
                if (needsWaw) words += " و ";
                if (number < 20)
                {
                    words += ones[number];
                }
                else
                {
                    words += tens[number / 10];
                    if ((number % 10) > 0)
                    {
                        words += " و " + ones[number % 10];
                    }
                }
            }

            return words;
        }
    }
    #endregion

    [HttpGet("export")]
    public async Task<IActionResult> ExportToExcel(
            [FromQuery] string? memberCode = null,
            [FromQuery] string? receiptId = null,
            [FromQuery] DateTime? startDate = null, [FromQuery] DateTime? endDate = null)
    {
        try
        {
            // --- 1. Build SQL Query ---
            string query = "SELECT * FROM ServiceEditView";
            var parameters = new List<object>();
            int parameterCounter = 0;

            if (string.IsNullOrWhiteSpace(memberCode) && string.IsNullOrWhiteSpace(receiptId) && !startDate.HasValue && !endDate.HasValue)
            {
                query = "SELECT TOP 30000 * FROM ServiceEditView ORDER BY ReceiptID DESC";
            }
            else
            {
                var conditions = new List<string>();
                if (!string.IsNullOrWhiteSpace(memberCode))
                {
                    conditions.Add($"MemberCode = @p{parameterCounter++}");
                    parameters.Add(memberCode);
                }
                if (!string.IsNullOrWhiteSpace(receiptId))
                {
                    conditions.Add($"ReceiptID = @p{parameterCounter++}");
                    parameters.Add(receiptId);
                }
                if (startDate.HasValue && endDate.HasValue)
                {
                    conditions.Add($"AddDate BETWEEN @p{parameterCounter++} AND @p{parameterCounter++}");
                    parameters.Add(startDate.Value);
                    parameters.Add(endDate.Value.Date.AddDays(1).AddMilliseconds(-1));
                }
                else if (startDate.HasValue)
                {
                    conditions.Add($"AddDate >= @p{parameterCounter++}");
                    parameters.Add(startDate.Value);
                }
                else if (endDate.HasValue)
                {
                    conditions.Add($"AddDate <= @p{parameterCounter++}");
                    parameters.Add(endDate.Value.Date.AddDays(1).AddMilliseconds(-1));
                }

                if (conditions.Count > 0)
                {
                    query += " WHERE " + string.Join(" AND ", conditions);
                }
            }

            // --- 2. Fetch Data into DataTable ---
            var dataTable = new DataTable();
            using (var connection = _context.Database.GetDbConnection())
            {
                await connection.OpenAsync();
                using (var command = connection.CreateCommand())
                {
                    command.CommandText = query;
                    for (int i = 0; i < parameters.Count; i++)
                    {
                        var parameter = command.CreateParameter();
                        parameter.ParameterName = $"@p{i}";
                        parameter.Value = parameters[i] ?? DBNull.Value;
                        command.Parameters.Add(parameter);
                    }
                    using (var reader = await command.ExecuteReaderAsync())
                    {
                        dataTable.Load(reader);
                    }
                }
            }

            // --- 3. Sort and Prepare DataTable ---
            if (dataTable.Rows.Count > 0)
            {
                var sortedDataTable = dataTable.AsEnumerable()
                    .OrderBy(row => row.Field<long>("ReceiptID"))
                    .CopyToDataTable();
                dataTable = sortedDataTable;
            }
            else
            {
                using var emptyWorkbook = new XLWorkbook();
                var emptySheet = emptyWorkbook.Worksheets.Add("NoData");
                emptySheet.Cell(1, 1).Value = "لا توجد بيانات للفترة المحددة.";
                using var emptyMemoryStream = new MemoryStream();
                emptyWorkbook.SaveAs(emptyMemoryStream);
                return File(emptyMemoryStream.ToArray(),
                           "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           "EmptyReport.xlsx");
            }

            // Rename columns to Arabic.
            if (dataTable.Columns.Contains("MemberCode")) dataTable.Columns["MemberCode"].ColumnName = "رقم القيد";
            if (dataTable.Columns.Contains("ReceiptID")) dataTable.Columns["ReceiptID"].ColumnName = "رقم الايصال";
            if (dataTable.Columns.Contains("MemberName")) dataTable.Columns["MemberName"].ColumnName = "اسم العضو";
            if (dataTable.Columns.Contains("Adddate")) dataTable.Columns["Adddate"].ColumnName = "التاريخ";
            if (dataTable.Columns.Contains("StringAmount")) dataTable.Columns["StringAmount"].ColumnName = "القيمة فقط وقدرها";
            if (dataTable.Columns.Contains("BranchName")) dataTable.Columns["BranchName"].ColumnName = "الفرع";
            if (dataTable.Columns.Contains("GovernrateName")) dataTable.Columns["GovernrateName"].ColumnName = "المحافظة";
            if (dataTable.Columns.Contains("cancelled")) dataTable.Columns["cancelled"].ColumnName = "لاغى";
            int nextYear = DateTime.Now.Year + 1;
            if (dataTable.Columns.Contains("_NextYearCalculation")) dataTable.Columns["_NextYearCalculation"].ColumnName = $"_اشتراك مقدم سنة{nextYear}";
            int nextNextYear = DateTime.Now.Year + 2;
            if (dataTable.Columns.Contains("_NextNextYearCalculation")) dataTable.Columns["_NextNextYearCalculation"].ColumnName = $"_اشتراك مقدم سنة{nextNextYear}";

            foreach (DataColumn column in dataTable.Columns)
            {
                string originalName = column.ColumnName;
                string sanitizedName = originalName.Replace("\n", " ").Trim();
                if (originalName != sanitizedName)
                {
                    column.ColumnName = sanitizedName;
                }
            }

            // --- 4. Create Excel Workbook and Sheets ---
            using var workbook = new XLWorkbook();

            var rawDataSheet = workbook.Worksheets.Add("RawData");
            rawDataSheet.RightToLeft = true;
            rawDataSheet.Cell(1, 1).InsertTable(dataTable.AsEnumerable());
            rawDataSheet.Columns().AdjustToContents();

            var reportSheet = workbook.Worksheets.Add("PivotReport");
            reportSheet.RightToLeft = true;

            reportSheet.Cell(1, 1).Value = "النقابه العامه للعلاج الطبيعى";
            reportSheet.Cell(1, 1).Style.Font.SetBold(true).Font.SetFontSize(14).Alignment.SetHorizontal(XLAlignmentHorizontalValues.Right);
            reportSheet.Range(1, 1, 1, 10).Merge();
            string branchName = (dataTable.Rows.Count > 0 && dataTable.Columns.Contains("الفرع")) ? dataTable.Rows[0]["الفرع"]?.ToString() ?? "" : "";
            reportSheet.Cell(2, 1).Value = $"فرع/ {branchName}";
            reportSheet.Cell(2, 1).Style.Font.SetBold(true).Alignment.SetHorizontal(XLAlignmentHorizontalValues.Right);
            reportSheet.Range(2, 1, 2, 10).Merge();
            string dateRangeText = "كشف حركه المتحصلات النقديه عن الفتره من ";
            if (startDate.HasValue && endDate.HasValue) dateRangeText += $"{startDate.Value:dd/MM/yyyy} الى {endDate.Value:dd/MM/yyyy}";
            else dateRangeText += "__________________ الى __________________";
            reportSheet.Cell(3, 1).Value = dateRangeText;
            reportSheet.Cell(3, 1).Style.Font.SetBold(true).Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
            reportSheet.Range(3, 1, 3, 10).Merge();

            // --- 6. Build Report Table Headers ---
            int headerRow = 5;
            reportSheet.Cell(headerRow, 1).Value = "الفرع";
            reportSheet.Cell(headerRow, 2).Value = "رقم القيد";
            reportSheet.Cell(headerRow, 3).Value = "رقم الايصال";
            reportSheet.Cell(headerRow, 4).Value = "اسم العضو";
            reportSheet.Cell(headerRow, 5).Value = "التاريخ";
            reportSheet.Cell(headerRow, 6).Value = "المحافظة";
            reportSheet.Cell(headerRow, 7).Value = "القيمة فقط وقدرها";
            reportSheet.Cell(headerRow, 8).Value = "لاغى";

            int colIndex = 9;
            var serviceColumns = new Dictionary<string, int>();
            foreach (DataColumn column in dataTable.Columns)
            {
                if (column.ColumnName.StartsWith("_") && dataTable.AsEnumerable().Any(r => r[column] != DBNull.Value && Convert.ToDouble(r[column]) > 0))
                {
                    string columnName = column.ColumnName.Replace("\n", " ").Trim();
                    reportSheet.Cell(headerRow, colIndex).Value = columnName;
                    serviceColumns.Add(column.ColumnName, colIndex++);
                }
            }
            int totalColIndex = colIndex;
            reportSheet.Cell(headerRow, totalColIndex).Value = "الإجمالى";
            reportSheet.Range(headerRow, 1, headerRow, totalColIndex).Style.Font.SetBold(true).Fill.SetBackgroundColor(XLColor.LightGray);

            // --- 7. Populate Report Data with Tafkeet ---
            int currentRow = headerRow + 1;
            foreach (DataRow dataRow in dataTable.Rows)
            {
                bool isCancelled = dataRow["لاغى"] != DBNull.Value && Convert.ToInt32(dataRow["لاغى"]) != 0;

                // Calculate total for the row (including cancelled)
                double rowTotal = 0;
                foreach (var service in serviceColumns.Keys)
                {
                    if (dataRow[service] != DBNull.Value && double.TryParse(dataRow[service].ToString(), out double value))
                    {
                        rowTotal += value;
                    }
                }

                // Populate basic columns
                reportSheet.Cell(currentRow, 1).Value = dataRow["الفرع"]?.ToString();
                reportSheet.Cell(currentRow, 2).Value = dataRow["رقم القيد"]?.ToString();
                reportSheet.Cell(currentRow, 3).Value = dataRow["رقم الايصال"]?.ToString();
                reportSheet.Cell(currentRow, 4).Value = dataRow["اسم العضو"]?.ToString();
                if (dataRow["التاريخ"] is DateTime dateValue)
                {
                    reportSheet.Cell(currentRow, 5).Value = dateValue.ToString("dd/MM/yyyy");
                }
                reportSheet.Cell(currentRow, 6).Value = dataRow["المحافظة"]?.ToString();

                // Use Tafkeet and add "(ملغى)" for cancelled receipts
                string amountInWords = TafkeetHelper.ToWords(Convert.ToDecimal(rowTotal));
                string amountText = isCancelled
                    ? $"فقط {amountInWords} جنيه مصري لا غير (ملغى)"
                    : $"فقط {amountInWords} جنيه مصري لا غير";
                reportSheet.Cell(currentRow, 7).Value = amountText;

                // Populate cancelled status
                reportSheet.Cell(currentRow, 8).Value = isCancelled ? "نعم" : "لا";

                // Populate service columns with actual values (even for cancelled)
                foreach (var service in serviceColumns)
                {
                    if (dataRow[service.Key] != DBNull.Value && double.TryParse(dataRow[service.Key].ToString(), out double value))
                    {
                        reportSheet.Cell(currentRow, service.Value).Value = value;
                    }
                    else
                    {
                        reportSheet.Cell(currentRow, service.Value).Value = 0;
                    }
                }

                // Populate the total column with actual total
                reportSheet.Cell(currentRow, totalColIndex).Value = rowTotal;

                // Apply formatting for cancelled rows
                if (isCancelled)
                {
                    reportSheet.Range(currentRow, 1, currentRow, totalColIndex).Style.Fill.SetBackgroundColor(XLColor.LightPink);
                    reportSheet.Cell(currentRow, 8).Style.Font.SetFontColor(XLColor.Red);
                }

                currentRow++;
            }

            // --- 8. Add Summary Row (only non-cancelled) ---
            int summaryRow = currentRow + 1;
            reportSheet.Cell(summaryRow, 1).Value = "الإجمالى";
            reportSheet.Cell(summaryRow, 1).Style.Font.SetBold(true);
            reportSheet.Range(summaryRow, 1, summaryRow, 8).Merge().Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);

            // Get column letters for criteria range (cancellation column)
            string criteriaColLetter = reportSheet.Column(8).ColumnLetter(); // Column H (لاغى)
            string criteriaRange = $"{criteriaColLetter}{headerRow + 1}:{criteriaColLetter}{currentRow - 1}";

            // Add SUMIFS formulas for numeric columns
            for (int i = 9; i <= totalColIndex; i++)
            {
                string colLetter = reportSheet.Column(i).ColumnLetter();
                string dataRange = $"{colLetter}{headerRow + 1}:{colLetter}{currentRow - 1}";

                // SUMIFS to include only non-cancelled rows
                reportSheet.Cell(summaryRow, i).FormulaA1 = $"SUMIFS({dataRange}, {criteriaRange}, \"لا\")";
                reportSheet.Cell(summaryRow, i).Style.Font.SetBold(true);
            }

            // --- 9. Finalize and Return File ---
            reportSheet.Columns().AdjustToContents();

            using var reportMemoryStream = new MemoryStream();
            workbook.SaveAs(reportMemoryStream);
            var content = reportMemoryStream.ToArray();

            return File(content,
                       "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                       $"Report_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error Details: {ex.Message}");
            Console.WriteLine($"Stack Trace: {ex.StackTrace}");
            return StatusCode(500, $"An error occurred during export: {ex.Message}");
        }
    }

    [HttpGet("DownloadReport")]
    public IActionResult DownloadReport()
    {
        var filePath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "DownloadReport.html");
        return PhysicalFile(filePath, "text/html");
    }

    [HttpGet("DownloadReport.html")]
    public IActionResult DownloadHtmlPage()
    {
        return RedirectToAction("DownloadReport");
    }
}