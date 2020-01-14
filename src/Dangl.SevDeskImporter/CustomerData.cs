using System.Collections.Generic;

namespace Dangl.SevDeskImporter
{
    public class CustomerData
    {
        public string CustomerNumber { get; set; }
        public string Name { get; set; }
        public string Website { get; set; }
        public string Notes { get; set; }
        public Address Address { get; set; }
        public List<CustomerContact> Contacts { get; set; }
        public bool IsActiveCustomer { get; set; }
    }
}
