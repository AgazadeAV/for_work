using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;
using System.IO;

class Program
{
    static void Main()
    {
        // Создаем новый Excel пакет
        using (var package = new ExcelPackage(new FileInfo("min_max_rates_calendar.xlsx")))
        {
            // Добавляем новый лист
            var worksheet = package.Workbook.Worksheets.Add("Data");

            // Добавляем заголовки столбцов
            worksheet.Cells["A1"].Value = "Month";
            worksheet.Cells["B1"].Value = "Min_Rate";
            worksheet.Cells["C1"].Value = "Max_Rate";

            // Добавляем данные
            string[] months = { "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec" };
            int[] minRates = { 10, 12, 15, 11, 13, 14, 16, 18, 20, 19, 17, 16 };
            int[] maxRates = { 20, 22, 25, 21, 23, 24, 26, 28, 30, 29, 27, 26 };

            for (int i = 0; i < months.Length; i++)
            {
                worksheet.Cells[i + 2, 1].Value = months[i];
                worksheet.Cells[i + 2, 2].Value = minRates[i];
                worksheet.Cells[i + 2, 3].Value = maxRates[i];
            }

            // Создаем диаграмму
            var barChart = worksheet.Drawings.AddChart("chart", eChartType.ColumnClustered) as ExcelBarChart;
            barChart!.Title.Text = "Min-Max Rates by Months";
            barChart.SetPosition(1, 0, 3, 0);
            barChart.SetSize(800, 600);
            barChart.Series.Add(ExcelRange.GetAddress(2, 2, 13, 2), ExcelRange.GetAddress(2, 1, 13, 1));
            barChart.Series.Add(ExcelRange.GetAddress(2, 3, 13, 3), ExcelRange.GetAddress(2, 1, 13, 1));
            barChart.Series[0].Header = "Min Rate";
            barChart.Series[1].Header = "Max Rate";

            // Сохраняем файл
            package.Save();
        }
    }
}