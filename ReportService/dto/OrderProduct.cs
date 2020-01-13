using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace ReportService.dto
{
    public class OrderProduct
    {
        public virtual string id { get; set; }

        public virtual string name { get; set; }

        public virtual string unitOfMeasure { get; set; }

        public virtual decimal unitPrice { get; set; }

        public virtual decimal boxWeight { get; set; }
    }
}
