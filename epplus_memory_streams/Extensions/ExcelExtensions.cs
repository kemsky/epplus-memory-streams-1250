using System.Drawing;
using ColorHelper;
using OfficeOpenXml;
using OfficeOpenXml.DataValidation;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Style;
using OfficeOpenXml.Table.PivotTable;
using ColorConverter = ColorHelper.ColorConverter;

namespace epplus_memory_streams.Extensions;

public static class ExcelExtensions
{
    private static readonly (int R, int G, int B)[] Colors =
    [
        (94, 185, 243),
        (32, 162, 103),
        (32, 100, 183),
        (48, 15, 60),
        (208, 59, 99),
        (176, 139, 192),
        (47, 48, 152),
        (94, 247, 93),
        (169, 76, 173),
        (84, 143, 245),
        (188, 200, 118),
        (115, 29, 229),
        (209, 231, 25),
        (31, 92, 97),
        (218, 185, 181),
        (13, 149, 4),
        (92, 76, 169),
        (222, 71, 72),
        (213, 80, 250),
        (61, 31, 186),
        (143, 218, 110),
        (205, 72, 110),
        (27, 23, 91),
        (106, 94, 16),
        (49, 232, 104),
        (78, 170, 70),
        (155, 78, 215),
        (16, 220, 57),
        (3, 119, 11),
        (203, 129, 43),
        (60, 65, 236),
        (163, 120, 148),
        (43, 179, 31),
        (5, 42, 199),
        (187, 4, 126),
        (177, 51, 81),
        (240, 59, 71),
        (161, 115, 147),
        (154, 132, 41),
        (250, 207, 51),
        (21, 94, 207),
        (156, 147, 74),
        (175, 147, 126),
        (31, 119, 252),
        (149, 160, 82),
        (38, 4, 239),
        (46, 195, 71),
        (228, 184, 225),
        (235, 184, 199),
        (144, 137, 249),
        (211, 25, 30),
        (68, 223, 196),
        (17, 203, 147)
    ];

    public const string CellPosition = "INDIRECT(\"RC\",FALSE)";

    private const string CurrencyFormatString = "$ #,##0.00";

    private const string PercentFormatString = "#,##0.00%";

    private const string NumberFormatString = "#,##0.00";

    private const string DateFormatString = "mm/dd/yyyy";

    private const string DateTimeFormatString = "mm/dd/yyyy hh:mm";

    private const string IntegerFormatString = "0";

    private const string DateMonthFormatString = "mmmm";

    private const string TextFormatString = "@";

    #region Core

    public static byte[] GetAsByteArrayReadable(this ExcelPackage package, string fileName)
    {
        return package.GetAsByteArray();
    }

    public static MemoryStream GetAsMemoryStreamReadable(this ExcelPackage package, string fileName)
    {
        package.Save();

        return (MemoryStream)package.Stream;
    }

    public static ExcelPackage SetTitle(this ExcelPackage package, string title)
    {
        package.Workbook.Properties.Title = title;
        return package;
    }

    public static ExcelPackage SetAuthor(this ExcelPackage package, string author)
    {
        package.Workbook.Properties.Author = author;
        return package;
    }

    public static ExcelPackage SetCreationDate(this ExcelPackage package, DateTime date)
    {
        package.Workbook.Properties.Created = date;
        return package;
    }

    public static ExcelPackage SetModificationDate(this ExcelPackage package, DateTime date)
    {
        package.Workbook.Properties.Modified = date;
        return package;
    }

    public static ExcelPackage SetCompany(this ExcelPackage package, string company = "Veritas Legal Plan")
    {
        package.Workbook.Properties.Company = company;
        return package;
    }

    public static ExcelWorksheet AddSheet(this ExcelPackage package, string name)
    {
        var sheetName = name;

        // max-length of sheet name
        if (sheetName.Length > 31)
        {
            sheetName = sheetName.Substring(0, 31 - 3) + "...";
        }

        var worksheet = package.Workbook.Worksheets.Add(sheetName);

        worksheet.DefaultColWidth = 25;

        return worksheet;
    }

    public static ExcelWorksheet AddSheet(this ExcelPackage package, string name, ExcelWorksheet copy)
    {
        var sheetName = name;

        // max-length of sheet name
        if (sheetName.Length > 31)
        {
            sheetName = sheetName.Substring(0, 31 - 3) + "...";
        }

        return package.Workbook.Worksheets.Add(sheetName, copy);
    }

    public static ExcelPackage AddSheets(this ExcelPackage package, ExcelWorksheets sources)
    {
        foreach (var source in sources)
        {
            var sheet = package.AddSheet(source.Name, source);

            sheet.Hidden = source.Hidden;
        }

        return package;
    }

    public static ExcelWorksheet Sheet(this ExcelPackage package, int index)
    {
        return package.Workbook.Worksheets[index];
    }

    public static ExcelRange CellRange(this ExcelWorksheet worksheet, int fromRow, int fromColumn, int toRow, int toColumn)
    {
        return worksheet.Cells[fromRow, fromColumn, toRow, toColumn];
    }

    public static ExcelRange CellRange(this ExcelWorksheet worksheet, string address)
    {
        return worksheet.Cells[address];
    }

    public static ExcelRange CellRange(this ExcelWorksheet worksheet)
    {
        return worksheet.Cells;
    }

    public static ExcelRange Cell(this ExcelWorksheet worksheet, int row, int column)
    {
        return worksheet.Cells[row, column];
    }

    public static ExcelRange Cell<TColumn>(this ExcelWorksheet worksheet, int row, TColumn column) where TColumn : struct, Enum
    {
        return worksheet.Cells[row, (int)(object)column];
    }

    public static ExcelWorksheet CalculateAndClearFormulas(this ExcelWorksheet worksheet)
    {
        worksheet.Cells.Calculate();

        foreach (var cell in worksheet.Cells.Where(x => !string.IsNullOrWhiteSpace(x.Formula)))
        {
            var value = cell.Value;
            cell.Formula = string.Empty;
            cell.Value = value;
        }

        return worksheet;
    }

    #endregion

    #region Style

    public static ExcelRange WrapText(this ExcelRange cell)
    {
        cell.Style.WrapText = true;
        return cell;
    }

    public static ExcelRange Bold(this ExcelRange cell)
    {
        cell.Style.Font.Bold = true;
        return cell;
    }

    public static ExcelRange Italic(this ExcelRange cell)
    {
        cell.Style.Font.Italic = true;
        return cell;
    }

    public static ExcelRange Size(this ExcelRange cell, float size)
    {
        cell.Style.Font.Size = size;
        return cell;
    }

    public static ExcelRange Alignment(this ExcelRange cell, ExcelVerticalAlignment alignment)
    {
        cell.Style.VerticalAlignment = alignment;
        return cell;
    }

    public static ExcelRange Alignment(this ExcelRange cell, ExcelHorizontalAlignment alignment)
    {
        cell.Style.HorizontalAlignment = alignment;
        return cell;
    }

    public static ExcelRange NumberFormat(this ExcelRange cell, string format = IntegerFormatString)
    {
        cell.Style.Numberformat.Format = format;
        return cell;
    }

    public static ExcelRange CurrencyFormat(this ExcelRange cell, string format = CurrencyFormatString)
    {
        cell.Style.Numberformat.Format = format;
        return cell;
    }

    public static ExcelRange PercentFormat(this ExcelRange cell, string format = PercentFormatString)
    {
        cell.Style.Numberformat.Format = format;
        return cell;
    }

    public static ExcelRange DateFormat(this ExcelRange cell, string format = DateFormatString)
    {
        cell.Style.Numberformat.Format = format;
        return cell;
    }

    public static ExcelRange DateTimeFormat(this ExcelRange cell, string format = DateTimeFormatString)
    {
        cell.Style.Numberformat.Format = format;
        return cell;
    }

    public static ExcelRange DateMonthFormat(this ExcelRange cell, string format = DateMonthFormatString)
    {
        cell.Style.Numberformat.Format = format;
        return cell;
    }

    public static ExcelRange TextFormat(this ExcelRange cell, string format = TextFormatString)
    {
        cell.Style.Numberformat.Format = format;
        return cell;
    }

    public static ExcelRange TextColor(this ExcelRange cell, string htmlColor)
    {
        var rgb = ColorConverter.HexToRgb(new HEX(htmlColor));
        cell.Style.Font.Color.SetColor(255, rgb.R, rgb.G, rgb.B);
        return cell;
    }

    public static ExcelRange Background(this ExcelRange cell, string htmlColor, ExcelFillStyle fillStyle = ExcelFillStyle.Solid)
    {
        cell.Style.Fill.PatternType = fillStyle;
        var rgb = ColorConverter.HexToRgb(new HEX(htmlColor));
        cell.Style.Fill.BackgroundColor.SetColor(255, rgb.R, rgb.G, rgb.B);
        return cell;
    }

    public static ExcelRange HeaderBorder(this ExcelRange cell, ExcelBorderStyle style)
    {
        cell.Style.Border.BorderAround(style);
        return cell;
    }

    public static ExcelRange BodyBorder(this ExcelRange cell, ExcelBorderStyle style)
    {
        cell.Border(bottom: style, right: style, diagonal: style);
        return cell;
    }

    public static ExcelRange BorderAround(this ExcelRange cell, ExcelBorderStyle style)
    {
        cell.Style.Border.BorderAround(style);
        return cell;
    }

    public static ExcelRange Border(this ExcelRange cell, ExcelBorderStyle? top = null, ExcelBorderStyle? right = null, ExcelBorderStyle? bottom = null, ExcelBorderStyle? left = null, ExcelBorderStyle? diagonal = null,
        bool? diagonalUp = null, bool? diagonalDown = null)
    {
        if (top != null)
        {
            cell.Style.Border.Top.Style = top.Value;
        }

        if (right != null)
        {
            cell.Style.Border.Right.Style = right.Value;
        }

        if (bottom != null)
        {
            cell.Style.Border.Bottom.Style = bottom.Value;
        }

        if (left != null)
        {
            cell.Style.Border.Left.Style = left.Value;
        }

        if (diagonal != null)
        {
            cell.Style.Border.Diagonal.Style = diagonal.Value;
        }

        if (diagonalDown != null)
        {
            cell.Style.Border.DiagonalDown = diagonalDown.Value;
        }

        if (diagonalUp != null)
        {
            cell.Style.Border.DiagonalUp = diagonalUp.Value;
        }

        return cell;
    }

    public static ExcelRange Merge(this ExcelRange cell)
    {
        cell.Merge = true;
        return cell;
    }

    public static ExcelRange AutoFilter(this ExcelRange cell, bool value = true)
    {
        cell.AutoFilter = value;
        return cell;
    }

    public static ExcelRange AutoColumns(this ExcelRange cell)
    {
        cell.AutoFitColumns();
        return cell;
    }

    public static ExcelDrawingFillBasic SetColor(this ExcelDrawingFillBasic fill, int index)
    {
        var color = Colors[index - (int)Math.Floor((decimal)index / Colors.Length) * Colors.Length];

        fill.Color = Color.FromArgb(color.R, color.G, color.B);

        return fill;
    }

    public static ExcelDrawingFillBasic SetColor(this ExcelDrawingFillBasic fill, int r, int g, int b)
    {
        fill.Color = Color.FromArgb(r, g, b);

        return fill;
    }

    #endregion

    #region Columns

    public static ExcelColumn Background(this ExcelColumn column, string htmlColor, ExcelFillStyle fillStyle = ExcelFillStyle.Solid)
    {
        column.Style.Fill.PatternType = fillStyle;

        var rgb = ColorConverter.HexToRgb(new HEX(htmlColor));

        column.Style.Fill.BackgroundColor.SetColor(255, rgb.R, rgb.G, rgb.B);

        return column;
    }

    public static ExcelColumn Width(this ExcelColumn column, int width)
    {
        column.Width = width;

        return column;
    }

    public static ExcelColumn WrapText(this ExcelColumn column)
    {
        column.Style.WrapText = true;

        return column;
    }

    #endregion

    #region Values

    public static T GetValueEx<T>(this ExcelRange cell)
    {
        try
        {
            if (cell == null)
            {
                return default;
            }

            if (typeof(T) == typeof(bool?))
            {
                if (cell.Value == null)
                {
                    return (T)(object)null;
                }

                if (cell.Value is string boolString)
                {
                    if (string.IsNullOrWhiteSpace(boolString))
                    {
                        return (T)(object)null;
                    }

                    if (TryParseBool(boolString, out var result))
                    {
                        return (T)(object)result;
                    }
                }
            }

            if (typeof(T) == typeof(bool))
            {
                if (cell.Value == null)
                {
                    return (T)(object)false;
                }

                if (cell.Value is string boolString)
                {
                    if (string.IsNullOrWhiteSpace(boolString))
                    {
                        return (T)(object)false;
                    }

                    if (TryParseBool(boolString, out var result))
                    {
                        return (T)(object)result;
                    }
                }
            }

            return cell.GetValue<T>();
        }
        catch (Exception e)
        {
            throw new Exception($"Failed to parse {typeof(T).Name}: \"{cell?.Text}\"", e);
        }
    }

    public static bool TryGetValueEx<T>(this ExcelRange cell, out T value)
    {
        try
        {
            value = cell.GetValueEx<T>();

            return true;
        }
        catch
        {
            value = default;

#pragma warning disable ERP022
            return false;
#pragma warning restore ERP022
        }
    }

    private static bool TryParseBool(string value, out bool result)
    {
        value = value.Trim();

        if (string.Equals(value, "yes", StringComparison.OrdinalIgnoreCase) || string.Equals(value, "true", StringComparison.OrdinalIgnoreCase))
        {
            result = true;

            return true;
        }

        if (string.Equals(value, "no", StringComparison.OrdinalIgnoreCase) || string.Equals(value, "false", StringComparison.OrdinalIgnoreCase))
        {
            result = false;

            return true;
        }

        result = false;

        return false;
    }

    public static ExcelRange SetFormulaEx(this ExcelRange cell, string value, bool asSharedFormula = true)
    {
        cell.SetFormula(value, asSharedFormula);
        return cell;
    }

    public static ExcelRange SetBoolean(this ExcelRange cell, bool value)
    {
        cell.Value = value;
        return cell;
    }

    public static ExcelRange SetBoolean(this ExcelRange cell, bool? value)
    {
        cell.Value = value;
        return cell;
    }

    public static ExcelRange SetPercent(this ExcelRange cell, double value, string format = PercentFormatString)
    {
        cell.Value = value / 100;
        cell.Style.Numberformat.Format = format;
        return cell;
    }

    public static ExcelRange SetPercent(this ExcelRange cell, decimal value, string format = PercentFormatString)
    {
        cell.Value = (double)value / 100;
        cell.Style.Numberformat.Format = format;
        return cell;
    }

    public static ExcelRange SetPercent(this ExcelRange cell, int value, string format = PercentFormatString)
    {
        cell.Value = value / 100;
        cell.Style.Numberformat.Format = format;
        return cell;
    }

    public static ExcelRange SetCurrency(this ExcelRange cell, decimal value, string format = CurrencyFormatString)
    {
        cell.Value = value;
        cell.Style.Numberformat.Format = format;
        return cell;
    }

    public static ExcelRange SetCurrency(this ExcelRange cell, decimal? value, string format = CurrencyFormatString)
    {
        cell.Value = value ?? 0;
        cell.Style.Numberformat.Format = format;
        return cell;
    }

    public static ExcelRange SetDate(this ExcelRange cell, DateTime value, string format = DateFormatString)
    {
        cell.Value = value;
        cell.Style.Numberformat.Format = format;
        return cell;
    }

    public static ExcelRange SetDate(this ExcelRange cell, DateTime? value, string format = DateFormatString)
    {
        cell.Value = value;
        cell.Style.Numberformat.Format = format;
        return cell;
    }

    public static ExcelRange SetDateTime(this ExcelRange cell, DateTime value, string format = DateTimeFormatString)
    {
        cell.Value = value;
        cell.Style.Numberformat.Format = format;
        return cell;
    }

    public static ExcelRange SetDateTime(this ExcelRange cell, DateTime? value, string format = DateTimeFormatString)
    {
        cell.Value = value;
        cell.Style.Numberformat.Format = format;
        return cell;
    }

    public static ExcelRange SetText(this ExcelRange cell, string value)
    {
        cell.Value = value;
        return cell;
    }

    public static ExcelRange SetNumber(this ExcelRange cell, int value, string format = IntegerFormatString)
    {
        cell.Value = value;
        cell.Style.Numberformat.Format = format;
        return cell;
    }

    public static ExcelRange SetNumber(this ExcelRange cell, int? value, string format = IntegerFormatString)
    {
        cell.Value = value;
        cell.Style.Numberformat.Format = format;
        return cell;
    }

    public static ExcelRange SetNumber(this ExcelRange cell, long value, string format = IntegerFormatString)
    {
        cell.Value = value;
        cell.Style.Numberformat.Format = format;
        return cell;
    }

    public static ExcelRange SetNumber(this ExcelRange cell, long? value, string format = IntegerFormatString)
    {
        cell.Value = value;
        cell.Style.Numberformat.Format = format;
        return cell;
    }

    public static ExcelRange SetNumber(this ExcelRange cell, double value, string format = NumberFormatString)
    {
        cell.Value = value;
        cell.Style.Numberformat.Format = format;
        return cell;
    }

    public static ExcelRange SetNumber(this ExcelRange cell, double? value, string format = NumberFormatString)
    {
        cell.Value = value;
        cell.Style.Numberformat.Format = format;
        return cell;
    }

    #endregion

    #region PivotTables

    public static ExcelPivotTable PivotTable(this ExcelWorksheet package, int index)
    {
        return package.PivotTables[index];
    }

    public static ExcelPivotTable SetCacheSource(this ExcelPivotTable table, ExcelRange source)
    {
        table.CacheDefinition.SourceRange = source;
        return table;
    }

    public static ExcelPivotTableField Field(this ExcelPivotTable table, int index)
    {
        return table.Fields[index];
    }

    #endregion

    #region Validations

    public static ExcelRange AddValidationLogical(this ExcelRange cell, string errorMessage = null)
    {
        var validation = cell.DataValidation.AddCustomDataValidation();
        validation.ShowErrorMessage = true;
        validation.Operator = ExcelDataValidationOperator.equal;
        validation.Formula.ExcelFormula = $"OR(ISLOGICAL({CellPosition}),LOWER({CellPosition})=\"true\",LOWER({CellPosition})=\"yes\",LOWER({CellPosition})=\"false\",LOWER({CellPosition})=\"no\")";
        validation.Error = errorMessage ?? "Invalid value. Valid values: True, False, Yes, No";

        return cell;
    }

    public static ExcelRange AddValidationEmail(this ExcelRange cell, string errorMessage = null)
    {
        var validation = cell.DataValidation.AddCustomDataValidation();

        validation.ShowErrorMessage = true;
        validation.Operator = ExcelDataValidationOperator.equal;
        validation.Formula.ExcelFormula = $"ISNUMBER(MATCH(\"?*@?*.?*\", {CellPosition}, 0))";
        validation.Error = errorMessage ?? "Email address is invalid";

        return cell;
    }

    public static ExcelRange AddValidationMajorityAge(this ExcelRange cell, string errorMessage = null)
    {
        var validation = cell.DataValidation.AddCustomDataValidation();

        validation.ShowErrorMessage = true;
        validation.Operator = ExcelDataValidationOperator.equal;
        validation.Formula.ExcelFormula = $"DATE(YEAR(TODAY())-18, MONTH(TODAY()), DAY(TODAY())) > {CellPosition}";
        validation.Error = errorMessage ?? "Minimum allowed age is 18";

        return cell;
    }

    public static ExcelRange AddValidationSsn(this ExcelRange cell, string errorMessage = null)
    {
        var validation = cell.DataValidation.AddCustomDataValidation();

        validation.ShowErrorMessage = true;
        validation.Operator = ExcelDataValidationOperator.equal;
        validation.Formula.ExcelFormula = $"AND(ISNUMBER(VALUE(SUBSTITUTE({CellPosition},\"-\",\"\"))), LEN(SUBSTITUTE({CellPosition},\"-\",\"\")) <= 9)";
        validation.Error = errorMessage ?? "This field must contain 9 digits or less";

        return cell;
    }

    public static ExcelRange AddValidationCreditCard(this ExcelRange cell, string errorMessage = null)
    {
        var validation = cell.DataValidation.AddCustomDataValidation();

        validation.ShowErrorMessage = true;
        validation.Operator = ExcelDataValidationOperator.equal;
        validation.Formula.ExcelFormula = $"AND(ISNUMBER(VALUE({CellPosition})), AND(LEN({CellPosition}) >= 12, LEN({CellPosition}) <= 19))";
        validation.Error = errorMessage ?? "The field must be whole number between 12 and 19 digits";

        return cell;
    }

    public static ExcelRange AddValidationExpirationDate(this ExcelRange cell, string errorMessage = null)
    {
        var validation = cell.DataValidation.AddCustomDataValidation();

        const string dividerPosition = $"FIND(\"/\",{CellPosition})";

        validation.ShowErrorMessage = true;
        validation.Operator = ExcelDataValidationOperator.equal;
        validation.Formula.ExcelFormula = $"IFERROR(AND(VALUE(LEFT({CellPosition},{dividerPosition}-1))<=12,VALUE(RIGHT({CellPosition},LEN({CellPosition})-{dividerPosition}))<=50),FALSE)";
        validation.Error = errorMessage ?? "Expiration date value is invalid, the expected format is MM/YY";

        return cell;
    }

    public static ExcelRange AddValidationZip(this ExcelRange cell, string errorMessage = null)
    {
        var validation = cell.DataValidation.AddCustomDataValidation();

        const string dividerPosition = $"FIND(\"-\",{CellPosition})";
        const string leftRule = $"VALUE(LEFT({CellPosition},IFERROR({dividerPosition}-1,LEN({CellPosition}))))<=99999";
        const string rightRule = $"VALUE(IFERROR(VALUE(SUBSTITUTE(RIGHT({CellPosition},LEN({CellPosition})-{dividerPosition}),\"-\",\"|\")),\"0\"))<=9999";

        validation.ShowErrorMessage = true;
        validation.Operator = ExcelDataValidationOperator.equal;
        validation.Formula.ExcelFormula = $"AND({leftRule},{rightRule})";
        validation.Error = errorMessage ?? "Zip is invalid. This field must be a 5-digit code with optional 4-digit extension XXXXX-XXXX";

        return cell;
    }

    public static ExcelRange AddValidationInteger(this ExcelRange cell, int? min = null, int? max = null, string errorMessage = null)
    {
        if (min == null && max == null)
        {
            return cell;
        }

        var validation = cell.DataValidation.AddIntegerDataValidation();

        validation.ShowErrorMessage = true;

        if (min != null && max != null)
        {
            validation.Operator = ExcelDataValidationOperator.between;
            validation.Formula.Value = min;
            validation.Formula2.Value = max;
            validation.Error = errorMessage ?? $"The field must be whole number between {min:D} and {max:D}";
        }
        else if (min != null)
        {
            validation.Operator = ExcelDataValidationOperator.greaterThanOrEqual;
            validation.Formula.Value = min;
            validation.Error = errorMessage ?? $"The field must be whole number greater than {min:D}";
        }
        else
        {
            validation.Operator = ExcelDataValidationOperator.lessThanOrEqual;
            validation.Formula.Value = max;
            validation.Error = errorMessage ?? $"The field must be whole number less than {max:D}";
        }

        return cell;
    }

    public static ExcelRange AddValidationDecimal(this ExcelRange cell, double? min = null, double? max = null, string errorMessage = null)
    {
        if (min == null && max == null)
        {
            return cell;
        }

        var validation = cell.DataValidation.AddDecimalDataValidation();

        validation.ShowErrorMessage = true;

        if (min != null && max != null)
        {
            validation.Operator = ExcelDataValidationOperator.between;
            validation.Formula.Value = min;
            validation.Formula2.Value = max;
            validation.Error = errorMessage ?? $"The field must be whole number or decimal between {min:F} and {max:F}";
        }
        else if (min != null)
        {
            validation.Operator = ExcelDataValidationOperator.greaterThanOrEqual;
            validation.Formula.Value = min;
            validation.Error = errorMessage ?? $"The field must be whole number or decimal greater than {min:F}";
        }
        else
        {
            validation.Operator = ExcelDataValidationOperator.lessThanOrEqual;
            validation.Formula.Value = max;
            validation.Error = errorMessage ?? $"The field must be whole number or decimal less than {max:F}";
        }

        return cell;
    }

    public static ExcelRange AddValidationDate(this ExcelRange cell, string errorMessage = null)
    {
        cell.AddValidationDate(new DateTime(2013, 1, 1), new DateTime(2050, 1, 1), errorMessage);

        return cell;
    }

    public static ExcelRange AddValidationDate(this ExcelRange cell, DateTime? min, DateTime? max, string errorMessage = null)
    {
        var validation = cell.DataValidation.AddDateTimeDataValidation();

        validation.ShowErrorMessage = true;

        if (min != null && max != null)
        {
            validation.Operator = ExcelDataValidationOperator.between;
            validation.Formula.Value = min;
            validation.Formula2.Value = max;
            validation.Error = errorMessage ?? $"The field must be date between {min:d} and {max:d}";
        }
        else if (min != null)
        {
            validation.Operator = ExcelDataValidationOperator.greaterThanOrEqual;
            validation.Formula.Value = min;
            validation.Error = errorMessage ?? $"The field must be date greater than {min:d}";
        }
        else
        {
            validation.Operator = ExcelDataValidationOperator.lessThanOrEqual;
            validation.Formula.Value = max;
            validation.Error = errorMessage ?? $"The field must be date less than {max:d}";
        }

        return cell;
    }

    public static ExcelRange AddValidationTextLength(this ExcelRange cell, int? min = null, int? max = null, string errorMessage = null)
    {
        if (min == null && max == null)
        {
            return cell;
        }

        var validation = cell.DataValidation.AddTextLengthDataValidation();

        validation.ShowErrorMessage = true;

        if (min != null && max != null)
        {
            validation.Operator = ExcelDataValidationOperator.between;
            validation.Formula.Value = min;
            validation.Formula2.Value = max;
            validation.Error = errorMessage ?? $"The field text lenght must be between {min:D} and {max:D}";
        }
        else if (min != null)
        {
            validation.Operator = ExcelDataValidationOperator.greaterThanOrEqual;
            validation.Formula.Value = min;
            validation.Error = errorMessage ?? $"The field text lenght must be greater than {min:D}";
        }
        else
        {
            validation.Operator = ExcelDataValidationOperator.lessThanOrEqual;
            validation.Formula.Value = max;
            validation.Error = errorMessage ?? $"The field text lenght must be less than {max:D}";
        }

        return cell;
    }

    public static ExcelRange AddValidationList(
        this ExcelRange cell,
        List<string> values,
        ExcelWorksheet validationSheet,
        bool allowEmptyValue = false,
        string errorMessage = null
    )
    {
        if (values.Count == 0)
        {
            return cell;
        }

        if (allowEmptyValue)
        {
            values.Insert(0, " ");
        }

        var column = validationSheet.Dimension != null ? validationSheet.Dimension.Columns + 1 : 1;

        for (var row = 0; row < values.Count; row++)
        {
            validationSheet.Cell(row + 1, column).SetText(values[row]);
        }

        validationSheet.Column(column).AutoFit();

        var address = validationSheet.CellRange(1, column, values.Count, column).FullAddressAbsolute;
        var validation = cell.DataValidation.AddListDataValidation();

        validation.Formula.ExcelFormula = address;
        validation.ShowErrorMessage = true;
        validation.Error = errorMessage ?? "Invalid value";

        return cell;
    }

    public static ExcelRange AddValidationDependentList(
        this ExcelRange targetRange,
        ExcelRange dependentRange,
        int principalColumnIndex,
        int dependentColumnIndex,
        ExcelWorksheet validationSheet,
        Dictionary<string, List<string>> values,
        bool allowEmptyValue = false,
        string errorMessage = null
    )
    {
        if (values.Count == 0)
        {
            return targetRange;
        }

        var columnIndex = validationSheet.Dimension != null ? validationSheet.Dimension.Columns + 1 : 1;
        var startColumnIndex = columnIndex;
        var rowIndex = 1;

        foreach (var (principalName, dependentNames) in values)
        {
            validationSheet.Cell(rowIndex, columnIndex).SetText(principalName);
            rowIndex++;

            if (allowEmptyValue)
            {
                validationSheet.Cell(rowIndex, columnIndex).SetText(" ");
                rowIndex++;
            }

            foreach (var dependentName in dependentNames)
            {
                validationSheet.Cell(rowIndex, columnIndex).SetText(dependentName);
                rowIndex++;
            }

            validationSheet.Column(columnIndex).AutoFit();

            columnIndex++;
            rowIndex = 1;
        }

        var startAddress = validationSheet.CellRange(1, startColumnIndex, 1, startColumnIndex).FullAddressAbsolute;
        var fullAddress = validationSheet.CellRange(1, startColumnIndex, 1, columnIndex - 1).FullAddressAbsolute;

        var principalValidation = targetRange.DataValidation.AddListDataValidation();

        principalValidation.Formula.ExcelFormula = fullAddress;
        principalValidation.ShowErrorMessage = true;
        principalValidation.Error = errorMessage ?? "Invalid value";

        var dependentValidation = dependentRange.DataValidation.AddListDataValidation();
        var size = $"COUNTA(OFFSET({startAddress}, 1, MATCH(INDIRECT(\"RC[{principalColumnIndex - dependentColumnIndex}]\",FALSE), {fullAddress}, 0) - 1, 500))";

        dependentValidation.Formula.ExcelFormula = $"OFFSET({startAddress}, 1, MATCH(INDIRECT(\"RC[{principalColumnIndex - dependentColumnIndex}]\",FALSE), {fullAddress}, 0) - 1, {size}, 1)";
        dependentValidation.ShowErrorMessage = true;
        dependentValidation.Error = errorMessage ?? "Invalid value";

        return targetRange;
    }

    #endregion
}