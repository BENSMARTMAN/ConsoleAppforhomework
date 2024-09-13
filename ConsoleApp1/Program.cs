using System.ComponentModel;
using QuestPDF.Fluent;
using QuestPDF.Previewer;
using QuestPDF.Helpers;
using System.Diagnostics;
using QuestPDF.Infrastructure;

Document.Create(cotainer =>
{
    cotainer.Page(page =>
    {
        page.Header()
            .Background(Colors.Green.Lighten2)
            .AlignCenter()
            .Text("這是表頭")
            .FontFamily("Microsoft JhengHei");
        page.Content().Padding(16).Table(table =>
        {
            table.ColumnsDefinition(columns =>
            {
                columns.ConstantColumn(100);
                columns.RelativeColumn(1);
                columns.RelativeColumn(2);
            });
            table.Header(headers =>
            {
                for (int i = 0; i < 3; i++)
                {
                    table.Cell().BorderBottom(1).Text($"Header {i}");
                }
            });
            for (int i = 0; i < 1000; i++)
            {
                table.Cell().BorderBottom(1).BorderColor(Colors.Grey.Lighten2).Text($"Item {i}");
            }
        });
        //.Background(Colors.Blue.Lighten2)
        //.Text("這是內文")
        //.BackgroundColor(Colors.Yellow.Lighten2);
        page.Footer()
             .AlignCenter()
             .PaddingVertical(20)
             .Text(text =>
             {
                 text.CurrentPageNumber();
                 text.Span(" / ");
                 text.TotalPages();
             });


    });
})
    .ShowInPreviewer();
   //.ShowInPreviewer();
//.GeneratePdf("output.pdf");
Process.Start(new ProcessStartInfo()
{
    FileName = "output.pdf",
    UseShellExecute = true
});


