using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace UCClientWebService.Models
{
    public class ResponseID
    {
        public string id { get; set; }
        public bool valid { get; set; }
        public string message { get; set; }
        public int RemainingDays { get; set; }
        public string messagelog { get; set; }

}
}
