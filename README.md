# How to modify the cell style while importing the worksheet

This sample demonstrates how to modify the cell style while importing the worksheet.

In `Spreadsheet` control, you can modify the cell style of particular cell or entire sheet while importing the Excel workbook. `CellStyle` includes font settings, alignment settings, border settings and fill color settings, etc.

For example, you can modify the cell style for the entire worksheet while importing the workbook by modifying the required properties of `CellStyle` of the `UsedRange` in `WorkBookLoaded` event of `Spreadsheet` control as illustrated in the following code example.

``` csharp
spreadSheetControl.WorkbookLoaded += SpreadSheetControl_WorkbookLoaded; ;
 
//To change the cellstyle while importing the sheet.
private void SpreadSheetControl_WorkbookLoaded(object sender, WorkbookLoadedEventArgs args)
{
    foreach (IWorksheet sheet in args.Workbook.Worksheets)
    {
        if (sheet.UsedRange.LastRow > 0 && sheet.UsedRange.LastColumn > 0)
        {
            var cellStyle = sheet.UsedRange.CellStyle;
            cellStyle.Font.FontName = "Arial Black";
            cellStyle.Font.Bold = true;
            cellStyle.Font.Color = ExcelKnownColors.Violet;
            cellStyle.Font.Size = 12;
            cellStyle.HorizontalAlignment = ExcelHAlign.HAlignLeft;
            cellStyle.Borders.LineStyle = ExcelLineStyle.Double;
        }
    }
}
```