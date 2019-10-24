using Syncfusion.UI.Xaml.Grid.Utility;
using Syncfusion.UI.Xaml.Spreadsheet.Helpers;
using Syncfusion.Windows.Tools.Controls;
using Syncfusion.XlsIO;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace ModifyCellStyle
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : RibbonWindow
    {
        public MainWindow()
        {
            InitializeComponent();
            spreadSheetControl.Open(@"..\..\Data\Sample.xlsx");
            spreadSheetControl.WorkbookLoaded += SpreadSheetControl_WorkbookLoaded; ;
        
        }
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
    }
}
