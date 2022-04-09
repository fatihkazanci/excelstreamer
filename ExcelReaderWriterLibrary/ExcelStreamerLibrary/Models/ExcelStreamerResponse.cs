using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelStreamerLibrary.Models
{
    public class ExcelStreamerResponse
    {
        public bool IsSuccess { get; set; } = true;
        public Exception Exception { get; set; }
        public object Result { get; set; }

        public void Error(Exception exception)
        {
            this.Exception = exception;
            this.IsSuccess = false;
        }
        public void Error()
        {
            this.IsSuccess = false;
        }

    }
}
