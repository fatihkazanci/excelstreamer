using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelStreamerLibrary.Attributes
{
    public class ExcelStreamerSheetName : Attribute
    {
        public string SheetName { get; set; }
        public ExcelStreamerSheetName(string sheetName)
        {
            this.SheetName = sheetName;
        }
    }
}
