using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelStreamerLibrary.Models
{
    public class ExcelStreamerCountResponse
    {
        public List<ExcelStreamerWorkSheetCountResponse> Sheets { get; set; } = new();
        public int TotalSheet { get; set; } = 0;
    }
}
