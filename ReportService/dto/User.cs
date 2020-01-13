using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace ReportService.dto
{
    public class User
    {
        public const string FRIEND = "A";

        public virtual string id { get; set; }

        public virtual string fullName { get; set; }

        public virtual string role { get; set; }

        public virtual string phone { get; set; }

        public virtual string email { get; set; }

        public virtual string referralFullName { get; set; }
    }
}
