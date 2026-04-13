using System.IO;
using System.Text;
using ClosedXML.Excel;
using ADPhoneBook.Models;

namespace ADPhoneBook.Services;

public static class ExportService
{
    // ─── Excel ────────────────────────────────────────────────────────────────

    public static void ExportToExcel(IEnumerable<Employee> employees, string filePath)
    {
        using var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Телефонная книга");

        // Заголовки
        string[] headers = { "ФИО", "Отдел", "Должность", "Рабочий тел.", "Мобильный тел.", "Email" };
        for (int c = 0; c < headers.Length; c++)
        {
            var cell = ws.Cell(1, c + 1);
            cell.Value = headers[c];
            cell.Style.Font.Bold      = true;
            cell.Style.Fill.BackgroundColor = XLColor.FromHtml("#1A3C5E");
            cell.Style.Font.FontColor = XLColor.White;
            cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
        }

        // Данные
        int row = 2;
        foreach (var e in employees)
        {
            ws.Cell(row, 1).Value = e.DisplayName;
            ws.Cell(row, 2).Value = e.Department;
            ws.Cell(row, 3).Value = e.Title;
            ws.Cell(row, 4).Value = e.WorkPhone;
            ws.Cell(row, 5).Value = e.MobilePhone;
            ws.Cell(row, 6).Value = e.Email;

            if (row % 2 == 0)
            {
                ws.Range(row, 1, row, 6)
                  .Style.Fill.BackgroundColor = XLColor.FromHtml("#F0F4F8");
            }
            row++;
        }

        // Форматирование
        ws.Columns().AdjustToContents();
        ws.SheetView.FreezeRows(1);
        ws.RangeUsed()!.SetAutoFilter();

        wb.SaveAs(filePath);
    }

    // ─── CSV ──────────────────────────────────────────────────────────────────

    public static void ExportToCsv(IEnumerable<Employee> employees, string filePath)
    {
        var sb = new StringBuilder();
        sb.AppendLine("ФИО;Отдел;Должность;Рабочий тел.;Мобильный тел.;Email");

        foreach (var e in employees)
        {
            sb.AppendLine(string.Join(";",
                Escape(e.DisplayName),
                Escape(e.Department),
                Escape(e.Title),
                Escape(e.WorkPhone),
                Escape(e.MobilePhone),
                Escape(e.Email)));
        }

        File.WriteAllText(filePath, sb.ToString(), new UTF8Encoding(true));
    }

    private static string Escape(string s)
    {
        if (s.Contains(';') || s.Contains('"') || s.Contains('\n'))
            return $"\"{s.Replace("\"", "\"\"")}\"";
        return s;
    }
}
