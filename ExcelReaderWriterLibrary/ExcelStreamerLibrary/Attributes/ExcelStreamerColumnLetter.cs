using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelStreamerLibrary.Attributes
{
    public class ExcelStreamerColumnLetter : Attribute
    {
        public string ColumnLetterName { get; set; }
        public ExcelStreamerColumnLetter(string columnLetterName)
        {
            this.ColumnLetterName = columnLetterName;
        }
    }
}
