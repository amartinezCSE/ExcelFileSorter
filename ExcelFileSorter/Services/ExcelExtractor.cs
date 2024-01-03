using ClosedXML.Excel;
using ExcelFileSorter.Interfaces;

namespace ExcelFileSorter.Services 
{
    public class ExcelExtractor : IExcelExtractor
    {
        public async Task<Dictionary<string, List<object>>> ExtractDataFromExcel(Stream stream)
        {
            var result = new Dictionary<string, List<object>>();

            using (var workbook = new XLWorkbook(stream))
            {
                foreach (var worksheet in workbook.Worksheets)
                {
                    var sheetName = worksheet.Name;
                    var columnData = ExtractColumns(worksheet);

                    result.Add(sheetName, columnData);
                }
            }

            return result;
        }

        private List<object> ExtractColumns(IXLWorksheet worksheet)
        {
            // Implement logic to extract the columns you need from each worksheet
            // For example, you can iterate over rows and columns and store the data

            // Dummy logic: Extracting data from the first two columns
            var data = new List<object>();
            foreach (var row in worksheet.RowsUsed())
            {
                data.Add(new
                {
                    Column1 = row.Cell(1).Value.ToString(),
                    Column2 = row.Cell(2).Value.ToString()
                    // Add more columns as needed
                });
            }

            return data;
        }
    }
}
