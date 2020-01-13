using ReportService.exception;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace ReportService.dto
{
    public class ExtractionError
    {
        public ExtractionError(ExcelExtractionException ex)
        {
            this.type = ex.InnerException.GetType().ToString();
            this.message = ex.Message;
            this.rowIndex = ex.rowIndex;
            this.colIndex = ex.colIndex;
        }

        public virtual string type { get; }

        public virtual string message { get; }

        public virtual Int32 rowIndex { get; }

        public virtual Int32 colIndex { get; }
    }
}
