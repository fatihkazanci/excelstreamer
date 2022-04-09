using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelStreamerLibrary.Models
{
    public abstract class ExcelStreamerWorkSheetObject
    {
        public int? _RowNumber { get; set; }
        public string _SheetName { get; set; }
    }
}
