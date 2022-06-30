using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EmailReseiver.Models
{
    public class ImportData
    {
        public Int64 Id { get; set; }
        public string? OrgName { get; set; }
        public string? MOD { get; set; }
        public string INN { get; set; }
        public string OKPO { get; set; }
        public string FinancingItem { get; set; }
        public string? ProductName { get; set; }
        public string? MedForm { get; set; }
        public string? SeriaNum { get; set; }
        public string? MNN { get; set; }
        public string MKB { get; set; }
        public string RecSeria { get; set; }
        public string? RecNum { get; set; }
        public DateTime RecDate { get; set; }
        public DateTime OtpuskDate { get; set; }
        public decimal Quant { get; set; }
        public string? OkeiName { get; set; }
        public decimal Price { get; set; }
        public decimal PSum { get; set; }
        public string? LastName { get; set; }
        public string? Name { get; set; }
        public string? MidName { get; set; }
        public DateTime DateOB { get; set; }
        public string? SNILS { get; set; }
        public DateTime InsertDate { get; set; }



    }
}
