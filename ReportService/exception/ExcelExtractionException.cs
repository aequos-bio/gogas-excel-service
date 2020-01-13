using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace ReportService.exception
{
    public class ExcelExtractionException : Exception
    {
        public ExcelExtractionException(Exception ex, Int32 rowIndex, Int32 colIndex) : base(ex.Message, ex)
        {
            this.rowIndex = rowIndex;
            this.colIndex = colIndex;
        }

        public virtual Int32 rowIndex { get; }

        public virtual Int32 colIndex { get; }
    }
}
