using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office365.ReportingWebServiceClient;

namespace o365stepbystep.Services
{
    public class ReportingWebServiceApi
    {
        private ReportingContext _context = new ReportingContext();
        private IReportVisitor _visitor = new CustomConsoleReportVisitor();
        public ReportingStream _stream1;

        public ReportingWebServiceApi(string username, string password, string operation)
        {
            this._context.UserName = username;
            this._context.Password = password;
            this._context.SetLogger(new CustomConsoleLogger());
            this._stream1 = new ReportingStream(this._context, operation, "stream1");
            this._stream1.RetrieveData();
        }

        public ReportingWebServiceApi(
            string username,
            string password,
            string operation,
            DateTime fromDate,
            DateTime toDate)
        {
            this._context.FromDateTime = fromDate;
            this._context.ToDateTime = toDate;
            new ReportingWebServiceApi(username, password, operation);
            
        }

       
    }

    internal class CustomConsoleReportVisitor : IReportVisitor
    {
        public override void VisitBatchReport()
        {
            foreach (ReportObject report in this.reportObjectList)
            {
                VisitReport(report);
            }
        }
        public override void VisitReport(ReportObject record)
        {
          //  Console.WriteLine("Record: " + record.Date.ToString());
        }

    }

    internal class CustomConsoleLogger : ITraceLogger
    {
        public void LogError(string message)
        {
            Console.WriteLine(message);
        }

        public void LogInformation(string message)
        {
            Console.WriteLine(message);
        }
    }

}
