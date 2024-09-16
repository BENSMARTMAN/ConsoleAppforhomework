using System.Globalization;
using QuestPDF.Fluent;
using QuestPDF.Helpers;
using QuestPDF.Previewer;

class Program
{
    static void Main(string[] args)
    {
        var document = Document.Create(container =>
        {
            // 12個月，從1月到12月
            for (int pageNumber = 1; pageNumber <= 12; pageNumber++)
            {
                // 添加每月的行事曆頁面
                container.Page(page =>
                {
                    page.Margin(20);

                    // 設置字型
                    page.Header().AlignCenter()
                        .Text(GetMonthName(pageNumber)) // 顯示頁碼和月份
                        .FontSize(24).Bold()
                        .FontFamily("Microsoft JhengHei");

                    // 設置表格的邊框
                    page.Content().Padding(10).Table(table =>
                    {
                        // 設置列數
                        table.ColumnsDefinition(column =>
                        {
                            for (int i = 0; i < 7; i++)  // 7 列，代表一周的7天
                            {
                                column.RelativeColumn();
                            }
                        });

                        // 添加星期標題，並設置邊框
                        var daysOfWeek = CultureInfo.CurrentCulture.DateTimeFormat.AbbreviatedDayNames;
                        foreach (var day in daysOfWeek)
                        {
                            table.Cell().Border(1).AlignCenter().Text(day).Bold()
                                .FontFamily("Microsoft JhengHei")
                                ; // 星期標題邊框
                        }

                        // 填入當月的日期，並設置邊框
                        var daysInMonth = DateTime.DaysInMonth(2025, pageNumber);
                        var firstDayOfMonth = new DateTime(2025, pageNumber, 1);
                        var startDayOfWeek = (int)firstDayOfMonth.DayOfWeek;

                        // 前面空格填滿之前的空白
                        for (int i = 0; i < startDayOfWeek; i++)
                        {
                            table.Cell().Border(1).Text(""); // 空白儲存格邊框
                        }

                        // 填寫日期，並設置邊框
                        for (int day = 1; day <= daysInMonth; day++)
                        {
                            table.Cell().Border(1).AlignCenter().Text(day.ToString())
                                ; // 日期儲存格邊框

                            // 如果填滿了一周的欄位（7列），換行
                            if ((startDayOfWeek + day) % 7 == 0)
                            {
                                // 自動換行
                            }
                        }

                        // 如果最後一行不滿 7 列，填入空白儲存格
                        var totalCells = startDayOfWeek + daysInMonth;
                        var remainingCells = (7 - (totalCells % 7)) % 7;

                        for (int i = 0; i < remainingCells; i++)
                        {
                            table.Cell().Border(1).Text(""); // 末尾空白儲存格邊框
                        }
                    })
                    ; // 表格外層邊框

                    // 頁碼顯示在頁尾
                    page.Footer().AlignCenter().Text($"Page {pageNumber}")
                        .FontFamily("Microsoft JhengHei");
                });
            }
        });

        // 使用 ShowInPreviewer 預覽
        document.ShowInPreviewer();
    }

    // 獲取月份名稱
    static string GetMonthName(int month)
    {
        return CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(month);
    }
}