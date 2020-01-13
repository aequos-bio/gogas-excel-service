using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace ReportService.dto
{
    public class PriceListProduct
    {
        public virtual string externalId { get; set; }

        public virtual string name { get; set; }

        public virtual string supplierExternalId { get; set; }

        public virtual string supplierName { get; set; }

        public virtual string supplierProvince { get; set; }

        public virtual string category { get; set; }

        public virtual string unitOfMeasure { get; set; }

        public virtual decimal boxWeight { get; set; }

        public virtual decimal unitPrice { get; set; }

        public virtual String notes { get; set; }

        public virtual String frequency { get; set; }

        public virtual Boolean wholeBoxesOnly { get; set; }

        public virtual Decimal? multiple { get; set; }
    }
}
