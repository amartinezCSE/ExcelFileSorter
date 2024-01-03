namespace ExcelFileSorter.Interfaces
{
    public interface IExcelExtractor
    {
        Task<Dictionary<string, List<object>>> ExtractDataFromExcel(Stream stream);
    }
}