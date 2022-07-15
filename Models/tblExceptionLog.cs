using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EmailReseiver.Models
{
    public class tblExceptionLog
    {
        public Int64 Id { get; set; }
        public DateTime LoggedDate { get; set; } = DateTime.Now;
        public string? MsgFrom {  get; set; }
        public string? MsgSubj {  get; set; }
        public string? MsgHasAttachment {  get; set; }
              
    }
}
