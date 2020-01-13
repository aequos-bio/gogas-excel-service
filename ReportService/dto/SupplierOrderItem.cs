﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace ReportService.dto
{
    public class SupplierOrderItem
    {
        public virtual string productId { get; set; }

        public virtual decimal unitPrice { get; set; }

        public virtual decimal boxWeight { get; set; }

        public virtual decimal quantity { get; set; }
    }
}
