using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using Table = DocumentFormat.OpenXml.Wordprocessing.Table;

namespace LearnOpenXml.Utils;

public class Cleaner
{
    public static void RemoveEmptyTableRows(OpenXmlElement element)
    {
        var tables = element.Descendants<Table>();
        var rowsToRemove = tables
            .SelectMany(table => table.Elements<TableRow>())
            .SelectMany(rows => rows.Elements<TableCell>(),
                (tableRow, tableCell) => new
                {
                    tableRow,
                    tableCell
                })
            .Where(x => ShouldBeRemoved(x.tableCell))
            .Select(x => x.tableRow)
            .ToArray();

        foreach (var tableRow in rowsToRemove)
        {
            tableRow.Remove();
        }
    }

    private static bool ShouldBeRemoved(OpenXmlElement cell)
    {
        const string removeIfEmptyKey = "#{remove if empty}";
        var text = cell.InnerText;

        if (!text.Contains(removeIfEmptyKey))
        {
            return false;
        }

        Replacer.ReplaceKey(cell, removeIfEmptyKey, string.Empty, false);

        return string.IsNullOrWhiteSpace(cell.InnerText);
    }
}