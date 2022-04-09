using ExcelStreamerLibrary.Attributes;
using ExcelStreamerLibrary.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelReaderWriterLibrary
{
    public class ExampleExcelModel : ExcelStreamerObject
    {
        [ExcelStreamerSheetName("Yapılacaklar Listesi")]
        public List<ExampleExcelWorkSheetModel> ToDoList { get; set; }
    }
}
