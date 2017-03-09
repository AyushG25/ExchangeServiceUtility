using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExchangeServiceUtility
{
    class EmailTemplateModel
    {
        public EmailTemplateModel(string mailType)
        {
            if (mailType.ToLower() == "workflowfailures")
            {
                this.ErrorMessage = EmailTemplateCollection.WorkFlowFailuresErrorMessage;
                this.Heading = EmailTemplateCollection.WorkFlowFailuresTableHeading;
                this.PST = DateTime.Now.AddDays(-1) + " - " + DateTime.Now;
            }
            else if(mailType.ToLower() == "mailboxupdatealert")
            {
                this.ErrorMessage = EmailTemplateCollection.MailboxUpdateAlertErrorMessage;
                this.Heading = EmailTemplateCollection.MailboxUpdateAlertTableHeading;
                this.PST = DateTime.Now.ToString();
                this.FooterMsg = EmailTemplateCollection.MailboxUpdateAlertCallToActionMsg + EmailTemplateCollection.MailboxUpdateAlertEscalationPathMsg;
            }

        }
        public string ErrorMessage { get; set; }
        public string Heading { get; set; }
        public string PST { get; set; }
        public string Content { get; set; }
        public string FooterMsg { get; set; }
    }
}
