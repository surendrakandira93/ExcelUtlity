using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace ExcelUtility.Models
{
    public class CustomerViewModel
    {
        public Guid? ID { get; set; }

        public byte? SalutationMasterId { get; set; }

        public string Name { get; set; }

        public string Pan { get; set; }

        public string Email { get; set; }

        public string STDCode { get; set; }

        public string PhoneNo { get; set; }

        public string Mobile { get; set; }

        public string GSTIN { get; set; }

        public string PlaceOfSupply { get; set; }
    }
}