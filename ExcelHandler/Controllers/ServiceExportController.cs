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
                        words += " جنيها و ";
                    else
                        words += " و ";
                    words += ConvertIntegerToWords(fractionalValue) + " قرشاً";
                }
            }
            else if (!string.IsNullOrEmpty(words))
            {
                words += " جنيه";
            }
            return words.Trim();
        }

        private static string ConvertIntegerToWords(long number)
        {
            if (number == 0) return "";
            string[] ones = { "", "واحد", "اثنان", "ثلاثة", "أربعة", "خمسة", "ستة", "سبعة", "ثمانية", "تسعة",
                            "عشرة", "أحد عشر", "اثنا عشر", "ثلاثة عشر", "أربعة عشر", "خمسة عشر",
                            "ستة عشر", "سبعة عشر", "ثمانية عشر", "تسعة عشر" };
            string[] tens = { "", "", "عشرون", "ثلاثون", "أربعون", "خمسون", "ستون", "سبعون", "ثمانون", "تسعون" };
            string[] hundreds = { "", "مائة", "مائتان", "ثلاثمائة", "أربعمائة", "خمسمائة",
                                "ستمائة", "سبعمائة", "ثمانمائة", "تسعمائة" };
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
                    int onesPart = (int)(number % 10);
                    int tensPart = (int)(number / 10);
                    if (onesPart > 0)
                    {
                        words += ones[onesPart] + " و " + tens[tensPart];
                    }
                    else
                    {
                        words += tens[tensPart];
                    }
                }
            }
            return words;
        }
    }
    #endregion

    // Helper function to merge two columns into one
    //private static void MergeTwoColumns(DataTable dataTable, string columnName1, string columnName2, string newColumnName)
    //{
    //    bool hasColumn1 = dataTable.Columns.Contains(columnName1);
    //    bool hasColumn2 = dataTable.Columns.Contains(columnName2);
    //    if (!hasColumn1 && !hasColumn2)
    //    {
    //        // Neither column exists, nothing to merge
    //        return;
    //    }
    //    // Create new column if it doesn't exist
    //    if (!dataTable.Columns.Contains(newColumnName))
    //    {
    //        dataTable.Columns.Add(newColumnName, typeof(decimal));
    //    }
    //    // Merge values
    //    foreach (DataRow row in dataTable.Rows)
    //    {
    //        decimal value1 = hasColumn1 && row[columnName1] != DBNull.Value ? Convert.ToDecimal(row[columnName1]) : 0;
    //        decimal value2 = hasColumn2 && row[columnName2] != DBNull.Value ? Convert.ToDecimal(row[columnName2]) : 0;
    //        row[newColumnName] = value1 + value2;
    //    }
    //    // Remove old columns
    //    if (hasColumn1) dataTable.Columns.Remove(columnName1);
    //    if (hasColumn2) dataTable.Columns.Remove(columnName2);
    //}
    private static void MergeColumns(DataTable dataTable, string[] columnsToMerge, string newColumnName)
    {
        var existingColumns = columnsToMerge.Where(col => dataTable.Columns.Contains(col)).ToList();

        if (!existingColumns.Any())
        {
            return;
        }

        if (!dataTable.Columns.Contains(newColumnName))
        {
            dataTable.Columns.Add(newColumnName, typeof(decimal));
        }

        foreach (DataRow row in dataTable.Rows)
        {
            decimal total = 0;

            foreach (var colName in existingColumns)
            {
                if (row[colName] != DBNull.Value)
                {
                    total += Convert.ToDecimal(row[colName]);
                }
            }

            row[newColumnName] = total;
        }

        foreach (var colName in existingColumns)
        {
            dataTable.Columns.Remove(colName);
        }
    }

    // New function to create the summary section at the bottom of the report
    private void CreateSummarySection(IXLWorksheet reportSheet, int summaryRow, Dictionary<string, int> serviceColumns,
        int headerRow, int currentRow, int totalColIndex, DataTable dataTable)
    {
        // Define base columns for fund calculations (these will be divided between funds)
        string[] baseColumns = {
            "_استمارة",
            "_NextNextYearCalculation",
            $"_اشتراك مقدم سنة{DateTime.Now.Year + 1}",
            $"_اشتراك مقدم سنة{DateTime.Now.Year + 2}",
            "_تحويل_اخصائى",
            "_رسم_القيد",
            "_سنوات_سابقه",
            "_العام_الحالى"
        };

        // Calculate summary statistics
        int validCount = 0;
        int cancelledCount = 0;
        double validTotal = 0;
        double cancelledTotal = 0;

        foreach (DataRow row in dataTable.Rows)
        {
            bool isCancelled = false;
            if (dataTable.Columns.Contains("لاغى"))
            {
                string cancelStatus = row["لاغى"]?.ToString() ?? "";
                isCancelled = cancelStatus == "الغاء بعد الطباعة" || cancelStatus == "محذوف قبل الطباعة";
            }

            // Calculate row total
            double rowTotal = 0;
            foreach (var service in serviceColumns)
            {
                if (row[service.Key] != DBNull.Value && double.TryParse(row[service.Key].ToString(), out double val))
                {
                    rowTotal += val;
                }
            }

            if (isCancelled)
            {
                cancelledCount++;
                cancelledTotal += rowTotal;
            }
            else
            {
                validCount++;
                validTotal += rowTotal;
            }
        }

        // Add spacing after the main table
        int cancelledSummaryRow = summaryRow + 3;

        // SECTION 1: Receipt Status Summary - Header
        reportSheet.Cell(cancelledSummaryRow, 1).Value = "حالة الإلغاء";
        reportSheet.Cell(cancelledSummaryRow, 2).Value = "عدد الإيصالات";
        reportSheet.Cell(cancelledSummaryRow, 3).Value = "إجمالي القيمة";
        reportSheet.Range(cancelledSummaryRow, 1, cancelledSummaryRow, 3).Style.Font.SetBold(true);
        reportSheet.Range(cancelledSummaryRow, 1, cancelledSummaryRow, 3).Style.Fill.SetBackgroundColor(XLColor.LightGray);

        // Valid receipts
        reportSheet.Cell(cancelledSummaryRow + 1, 1).Value = "إيصالات صالحة";
        reportSheet.Cell(cancelledSummaryRow + 1, 2).Value = validCount;
        reportSheet.Cell(cancelledSummaryRow + 1, 3).Value = validTotal;

        // Cancelled receipts
        reportSheet.Cell(cancelledSummaryRow + 2, 1).Value = "إيصالات ملغاة";
        reportSheet.Cell(cancelledSummaryRow + 2, 2).Value = cancelledCount;
        reportSheet.Cell(cancelledSummaryRow + 2, 3).Value = cancelledTotal;
        reportSheet.Cell(cancelledSummaryRow + 2, 1).Style.Font.SetFontColor(XLColor.Red);
        reportSheet.Cell(cancelledSummaryRow + 2, 2).Style.Font.SetFontColor(XLColor.Red);
        reportSheet.Cell(cancelledSummaryRow + 2, 3).Style.Font.SetFontColor(XLColor.Red);

        // Total row
        reportSheet.Cell(cancelledSummaryRow + 3, 1).Value = "الإجمالي";
        reportSheet.Cell(cancelledSummaryRow + 3, 2).Value = validCount + cancelledCount;
        reportSheet.Cell(cancelledSummaryRow + 3, 3).Value = validTotal + cancelledTotal;

        // SECTION 2: Fund Calculations
        // Calculate Union Fund total (صندوق النقابة)
        double baseTotal = 0;
        foreach (var serviceKvp in serviceColumns)
        {
            string columnName = serviceKvp.Key;
            if (baseColumns.Any(bc => columnName.EndsWith(bc)))
            {
                // Get the column range from the main data table
                var colRange = reportSheet.Range(headerRow + 1, serviceKvp.Value, currentRow - 1, serviceKvp.Value);
                foreach (var cell in colRange.Cells())
                {
                    if (!cell.IsEmpty() && double.TryParse(cell.Value.ToString(), out double val))
                    {
                        baseTotal += val;
                    }
                }
            }
        }

        // Sum other service columns excluding base columns and pension columns
        double otherTotal = 0;
        foreach (var serviceKvp in serviceColumns)
        {
            string columnName = serviceKvp.Key;
            if (!baseColumns.Any(bc => columnName.EndsWith(bc)) && !columnName.Contains("معاشات"))
            {
                var colRange = reportSheet.Range(headerRow + 1, serviceKvp.Value, currentRow - 1, serviceKvp.Value);
                foreach (var cell in colRange.Cells())
                {
                    if (!cell.IsEmpty() && double.TryParse(cell.Value.ToString(), out double val))
                    {
                        otherTotal += val;
                    }
                }
            }
        }

        // Calculate Pension Fund total (صندوق المعاشات)
        double pensionsTotal = 0;
        foreach (var serviceKvp in serviceColumns)
        {
            string columnName = serviceKvp.Key;
            if (columnName.Contains("معاشات"))
            {
                var colRange = reportSheet.Range(headerRow + 1, serviceKvp.Value, currentRow - 1, serviceKvp.Value);
                foreach (var cell in colRange.Cells())
                {
                    if (!cell.IsEmpty() && double.TryParse(cell.Value.ToString(), out double val))
                    {
                        pensionsTotal += val;
                    }
                }
            }
        }

        // Union Fund calculation (صندوق النقابة)
        double unionFundTotal = (baseTotal / 2) + otherTotal;
        int unionFundRow = cancelledSummaryRow + 5;
        reportSheet.Cell(unionFundRow, 1).Value = "صندوق النقابه";
        reportSheet.Cell(unionFundRow, 1).Style.Font.SetBold(true);
        reportSheet.Range(unionFundRow, 1, unionFundRow, totalColIndex - 1).Merge();
        reportSheet.Cell(unionFundRow, totalColIndex).Value = unionFundTotal;
        reportSheet.Cell(unionFundRow, totalColIndex).Style.Font.SetBold(true);
        reportSheet.Range(unionFundRow, 1, unionFundRow, totalColIndex).Style.Border.SetOutsideBorder(XLBorderStyleValues.Medium);

        // Pension Fund calculation (صندوق المعاشات)
        double pensionFundTotal = (baseTotal / 2) + pensionsTotal;
        int pensionFundRow = unionFundRow + 1;
        reportSheet.Cell(pensionFundRow, 1).Value = "صندوق المعاشات";
        reportSheet.Cell(pensionFundRow, 1).Style.Font.SetBold(true);
        reportSheet.Range(pensionFundRow, 1, pensionFundRow, totalColIndex - 1).Merge();
        reportSheet.Cell(pensionFundRow, totalColIndex).Value = pensionFundTotal;
        reportSheet.Cell(pensionFundRow, totalColIndex).Style.Font.SetBold(true);
        reportSheet.Range(pensionFundRow, 1, pensionFundRow, totalColIndex).Style.Border.SetOutsideBorder(XLBorderStyleValues.Medium);

        // Format all rows for right-to-left display
        reportSheet.Range(cancelledSummaryRow, 1, pensionFundRow, totalColIndex).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Right);
    }

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
            // --- 2.5: Merge required columns ---
            MergeColumns(dataTable, new[] { "_كارنية", "_كارنيه" }, "كارنية");
            MergeColumns(dataTable, new[] { "_اعاده_تسجيل", "_رسم_القيد", "_رسوم_ممارس" }, "رسم القيد");
            MergeColumns(dataTable, new[] { "__انهاء_اجراءات", "_انهاء_اجراءات" }, "إنهاء إجراءات");
            MergeColumns(dataTable, new[] { "_مبلغ_لدفعة_النقدية_المتقدمة", "[_دمغة_علاج_طبيعى]" }, "دمغة علاج طبيعى");
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
            // --- 3.1 Process Cancellation Status ---
            // حل المشكلة: تغيير نوع العمود قبل تعبئة القيم النصية
            if (dataTable.Columns.Contains("cancelled") && dataTable.Columns.Contains("printed"))
            {
                // إنشاء عمود مؤقت من نوع السلسلة النصية
                var tempCancelledColumn = new DataColumn("tempCancelled", typeof(string));
                dataTable.Columns.Add(tempCancelledColumn);
                foreach (DataRow row in dataTable.Rows)
                {
                    object cancelledValue = row["cancelled"];
                    object printedValue = row["printed"];
                    bool isCancelled = (cancelledValue != DBNull.Value) && Convert.ToInt32(cancelledValue) == 1;
                    bool isPrinted = (printedValue != DBNull.Value) && Convert.ToInt32(printedValue) == 1;
                    if (isCancelled)
                    {
                        if (isPrinted)
                            row["tempCancelled"] = "الغاء بعد الطباعة";
                        else
                            row["tempCancelled"] = "محذوف قبل الطباعة";
                    }
                    else
                    {
                        row["tempCancelled"] = "لا";
                    }
                }
                // إزالة العمود الأصلي وإعادة تسمية العمود المؤقت
                dataTable.Columns.Remove("cancelled");
                dataTable.Columns["tempCancelled"].ColumnName = "cancelled";
            }
            dataTable.Columns.Remove("printed");
            // Rename columns to Arabic.
            if (dataTable.Columns.Contains("MemberCode")) dataTable.Columns["MemberCode"].ColumnName = "رقم القيد";
            if (dataTable.Columns.Contains("ReceiptID")) dataTable.Columns["ReceiptID"].ColumnName = "رقم الايصال";
            if (dataTable.Columns.Contains("MemberName")) dataTable.Columns["MemberName"].ColumnName = "اسم العضو";
            if (dataTable.Columns.Contains("Adddate")) dataTable.Columns["Adddate"].ColumnName = "التاريخ";
            if (dataTable.Columns.Contains("StringAmount")) dataTable.Columns["StringAmount"].ColumnName = "القيمة فقط وقدرها";
            if (dataTable.Columns.Contains("BranchName")) dataTable.Columns["BranchName"].ColumnName = "الفرع";
            if (dataTable.Columns.Contains("GovernrateName")) dataTable.Columns["GovernrateName"].ColumnName = "المحافظة";
            if (dataTable.Columns.Contains("cancelled")) dataTable.Columns["cancelled"].ColumnName = "لاغى";
            if (dataTable.Columns.Contains("DelUser")) dataTable.Columns["DelUser"].ColumnName = "الغاء بواسطة حساب";
            if (dataTable.Columns.Contains("DelDate")) dataTable.Columns["DelDate"].ColumnName = "تاريخ الالغاء";
            if (dataTable.Columns.Contains("FullName")) dataTable.Columns["FullName"].ColumnName = "الغاء بواسطة اسم";
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
            reportSheet.Range(1, 1, 1, 12).Merge();
            string branchName = (dataTable.Rows.Count > 0 && dataTable.Columns.Contains("الفرع")) ? dataTable.Rows[0]["الفرع"]?.ToString() ?? "" : "";
            reportSheet.Cell(2, 1).Value = $"فرع/ {branchName}";
            reportSheet.Cell(2, 1).Style.Font.SetBold(true).Alignment.SetHorizontal(XLAlignmentHorizontalValues.Right);
            reportSheet.Range(2, 1, 2, 12).Merge();
            string dateRangeText = "كشف حركه المتحصلات النقديه عن الفتره من ";
            if (startDate.HasValue && endDate.HasValue) dateRangeText += $"{startDate.Value:dd/MM/yyyy} الى {endDate.Value:dd/MM/yyyy}";
            else dateRangeText += "__________________ الى __________________";
            reportSheet.Cell(3, 1).Value = dateRangeText;
            reportSheet.Cell(3, 1).Style.Font.SetBold(true).Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
            reportSheet.Range(3, 1, 3, 12).Merge();
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
            reportSheet.Cell(headerRow, 9).Value = "الغاء بواسطة حساب";
            reportSheet.Cell(headerRow, 10).Value = "تاريخ الالغاء";
            reportSheet.Cell(headerRow, 11).Value = "الغاء بواسطة اسم";
            int colIndex = 12;
            var serviceColumns = new Dictionary<string, int>();
            // قائمة الأعمدة المدمجة التي يجب أن تظهر في التقرير
            var mergedColumnsToShow = new List<string> { "كارنية","رسم القيد", "إنهاء إجراءات", "دمغة علاج طبيعى" };
            foreach (DataColumn column in dataTable.Columns)
            {
                string columnName = column.ColumnName.Replace("\n", " ").Trim();
                bool shouldInclude = false;
                // تحقق إذا كان العمود من الأعمدة المدمجة أو يبدأ بشرطة سفلية
                if (mergedColumnsToShow.Contains(columnName)
                    || columnName.StartsWith("_"))
                {
                    // تحقق إذا كان العمود يحتوي على قيم موجبة
                    if (dataTable.AsEnumerable().Any(r =>
                        r[column] != DBNull.Value && Convert.ToDouble(r[column]) > 0))
                    {
                        shouldInclude = true;
                    }
                }
                if (shouldInclude)
                {
                    // إزالة الشرطة الأولى من الاسم للعرض
                    string displayName = columnName.StartsWith("_")
                        ? columnName.Substring(1)
                        : columnName;
                    reportSheet.Cell(headerRow, colIndex).Value = displayName;
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
                bool isCancelled = dataRow["لاغى"] != DBNull.Value &&
                                  (dataRow["لاغى"].ToString() == "الغاء بعد الطباعة" ||
                                   dataRow["لاغى"].ToString() == "محذوف قبل الطباعة");
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
                    ? $"فقط {amountInWords} مصري لا غير (ملغى)"
                    : $"فقط {amountInWords} مصري لا غير";
                reportSheet.Cell(currentRow, 7).Value = amountText;
                // Populate cancellation status
                reportSheet.Cell(currentRow, 8).Value = dataRow["لاغى"]?.ToString();
                // Populate cancellation details only if cancelled
                if (isCancelled)
                {
                    reportSheet.Cell(currentRow, 9).Value = dataRow["الغاء بواسطة حساب"]?.ToString();
                    if (dataRow["تاريخ الالغاء"] is DateTime delDate)
                    {
                        reportSheet.Cell(currentRow, 10).Value = delDate.ToString("dd/MM/yyyy");
                    }
                    else
                    {
                        reportSheet.Cell(currentRow, 10).Value = dataRow["تاريخ الالغاء"]?.ToString();
                    }
                    reportSheet.Cell(currentRow, 11).Value = dataRow["الغاء بواسطة اسم"]?.ToString();
                }
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
            reportSheet.Range(summaryRow, 1, summaryRow, 11).Merge().Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
            // Get column letters for criteria range (cancellation column)
            string criteriaColLetter = reportSheet.Column(8).ColumnLetter(); // Column H (لاغى)
            string criteriaRange = $"{criteriaColLetter}{headerRow + 1}:{criteriaColLetter}{currentRow - 1}";
            // Add SUMIFS formulas for numeric columns
            for (int i = 12; i <= totalColIndex; i++)
            {
                string colLetter = reportSheet.Column(i).ColumnLetter();
                string dataRange = $"{colLetter}{headerRow + 1}:{colLetter}{currentRow - 1}";
                // SUMIFS to include only non-cancelled rows
                reportSheet.Cell(summaryRow, i).FormulaA1 = $"SUMIFS({dataRange}, {criteriaRange}, \"لا\")";
                reportSheet.Cell(summaryRow, i).Style.Font.SetBold(true);
            }

            // --- 9. Add the detailed summary section with fund calculations ---
            CreateSummarySection(reportSheet, summaryRow, serviceColumns, headerRow, currentRow, totalColIndex, dataTable);

            // --- 10. Finalize and Return File ---
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