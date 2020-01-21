using BSJobBase;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CommissionsCreate
{
   public class ExcelFormatOption
    {
        public ExcelFormatOption()
        {
            //setting default values
            StyleName = null;
            MergeCells = false;
            IsBold = false;
            IsUnderLine = false;
            WrapText = false;
            BorderBottomLineStyle = 0; 
            BorderTopLineStyle = 0; 
            BorderLeftLineStyle = 0; 
            BorderRightLineStyle = 0; 
            HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
        }

        public string StyleName { get; set; }

        public bool MergeCells { get; set; }

        public bool IsBold { get; set; }

        public bool IsUnderLine { get; set; }

        public bool WrapText { get; set; }

        public Microsoft.Office.Interop.Excel.XlLineStyle BorderTopLineStyle { get; set; }

        public Microsoft.Office.Interop.Excel.XlLineStyle BorderBottomLineStyle { get; set; }

        public Microsoft.Office.Interop.Excel.XlLineStyle BorderLeftLineStyle { get; set; }

        public Microsoft.Office.Interop.Excel.XlLineStyle BorderRightLineStyle { get; set; }



        public bool CenterText { get; set; }

        public Microsoft.Office.Interop.Excel.XlHAlign HorizontalAlignment { get; set; }

        public string NumberFormat { get; set; }

        public ExcelColor FillColor { get; set; }

        public ExcelColor TextColor { get; set; }


    }


}
