using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace ReportService.dto
{
    public class OrderExportRequest
    {
        public virtual List<OrderProduct> products { get; set; }

        public virtual List<User> users { get; set; }

        public virtual List<OrderItem> userOrder { get; set; }

        public virtual List<SupplierOrderItem> supplierOrder { get; set; }

        public virtual Boolean friends { get; set; }
    }
}
