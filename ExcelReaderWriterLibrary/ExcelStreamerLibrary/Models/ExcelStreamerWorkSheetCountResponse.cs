using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelStreamerLibrary.Models
{
    public class ExcelStreamerWorkSheetCountResponse
    {
        public string SheetName { get; set; }
        public int RowCount { get; set; }
        public int ColumnCount { get; set; }
    }
}
