using DocumentFormat.OpenXml.Wordprocessing;

namespace DocumentProcessor.Services
{
    public interface IHtmlToWordConverter
    {
        string ConvertHtmlToWordFormat(string html);
        Table CreateTable(string[][] data);

        string ConvertListToWordFormat(string htmlList, int numId);
    }
}
