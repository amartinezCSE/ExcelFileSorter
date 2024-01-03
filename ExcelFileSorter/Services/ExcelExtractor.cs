using ClosedXML.Excel;

namespace ExcelFileSorter.Services 
{
    public static class ExcelExtractor
    {
        
        public static List<EngineCommandObject> MapFile(this Stream fileStream)
        {
            var worksheet = GetWorksheet(fileStream);
            var validHeaders = worksheet.GetWorksheetHeaders();
            var validRows = worksheet.GetContentRows();
            return validRows.GetObjects(validHeaders);
        }

        private static IXLWorksheet GetWorksheet(Stream fileStream)
        {
            var workbook = new XLWorkbook(fileStream);
            return workbook.Worksheets.First();
        }

        private static Dictionary<string, string> GetWorksheetHeaders(this IXLWorksheet worksheet)
        {
            var potentialHeaders = GetObjectHeaders();
            var firstRow = worksheet.Row(1).Cells(true);
            return firstRow
                .Where(x => x != null && potentialHeaders.ContainsKey(x.Value.ToString()?.ToLower() ?? string.Empty))
                .GroupBy(x => x.Value.ToString()?.ToLower() ?? string.Empty)
                .ToDictionary(x => x.Key, x => x.First().Address.ColumnLetter.ToString());
        }

        private static Dictionary<string, string> GetObjectHeaders()
        {
            var propertyNames = typeof(EngineCommandObject).GetProperties().Select(x => x.Name);
            return propertyNames.Distinct(StringComparer.OrdinalIgnoreCase).ToDictionary(x => x.ToLower());
        }

        private static IEnumerable<IXLRow> GetContentRows(this IXLWorksheet worksheet)
        {
            return worksheet.Rows().Where(x => x.RowNumber() != 1 && !x.IsEmpty());
        }

        private static List<EngineCommandObject> GetObjects(this IEnumerable<IXLRow> validRows,
            IReadOnlyDictionary<string, string> worksheetHeaders)
        {
            return validRows.Select(x => x.GetOneObject(worksheetHeaders)).ToList();
        }

        private static EngineCommandObject GetOneObject(this IXLRow row, IReadOnlyDictionary<string, string> worksheetHeaders)
        {
            return new EngineCommandObject
            {
                Model = row.GetPropertyStringValue(worksheetHeaders, "Model"),
                PerformanceNumber = row.GetPropertyStringValue(worksheetHeaders, "PerformanceNumber"),
                Power = row.GetPropertyDecimalValue(worksheetHeaders, "Power"),
                Speed = row.GetPropertyDecimalValue(worksheetHeaders, "Speed"),
                BrakeSpecificFuelConsumption = worksheetHeaders.ContainsKey("BSFC")
                        ? row.GetPropertyDecimalValue(worksheetHeaders, "BSFC")
                    : row.GetPropertyDecimalValue(worksheetHeaders, "BrakeSpecificFuelConsumption"),
                BrakeSpecificOilConsumption = worksheetHeaders.ContainsKey("BSOC")
                    ? row.GetPropertyDecimalValue(worksheetHeaders, "BSOC")
                    : row.GetPropertyDecimalValue(worksheetHeaders, "BrakeSpecificOilConsumption")
            };
        }

        private static string GetPropertyStringValue(this IXLRow row,
            IReadOnlyDictionary<string, string> worksheetHeaders, string rowKey)
        {
            var rowIsValid = worksheetHeaders.ContainsKey(rowKey.ToLower());
            return !rowIsValid ? string.Empty : row.Cell($"{worksheetHeaders[rowKey.ToLower()]}").Value.ToString() ?? string.Empty;
        }

        private static decimal GetPropertyDecimalValue(this IXLRow row, IReadOnlyDictionary<string, string> worksheetHeaders, string rowKey)
        {
            var rowIsValid = worksheetHeaders.ContainsKey(rowKey.ToLower());
            if (!rowIsValid) return 0m;

            return !decimal.TryParse(row.Cell($"{worksheetHeaders[rowKey.ToLower()]}").Value.ToString(),
                out var validDecimal) ? 0m : validDecimal;
        }
       
    }
}
