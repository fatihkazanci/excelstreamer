using ExcelStreamerLibrary.Attributes;
using ExcelStreamerLibrary.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelReaderWriterLibrary
{
    public class ExampleExcelWorkSheetModel: ExcelStreamerWorkSheetObject
    {
        [ExcelStreamerColumnLetter("a")]
        public string Name { get; set; }
        [ExcelStreamerColumnLetter("b")]
        public string Surname { get; set; }
    }
}
