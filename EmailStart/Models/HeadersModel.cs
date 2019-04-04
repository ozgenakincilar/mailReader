using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace EmailStart.Models
{
    public class HeadersModel
    {
        public string From { get; set; }
        public string FileName { get; set; }
        public DateTime SentDate { get; set; }  
        public string Subject { get; set; }
        public string Body { get; set; }
        public string MessageId { get; set; }

    }
    public class BelgeModel
    {
        public List<string> Kolon { get; set; }
    }
    
}