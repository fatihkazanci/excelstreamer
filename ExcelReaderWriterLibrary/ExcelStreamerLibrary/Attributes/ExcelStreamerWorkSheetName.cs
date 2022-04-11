using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelStreamerLibrary.Attributes
{
    public class ExcelStreamerWorkSheetName : Attribute
    {
        public string SheetName { get; set; }
        public ExcelStreamerWorkSheetName(string sheetName)
        {
            this.SheetName = sheetName;
        }
    }
}
