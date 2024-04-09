using Microsoft.AspNetCore.Mvc;
using Microsoft.Reporting.NETCore;
using MyReports;
using RDLC_Demo.Models;
using System.Diagnostics;
using System.Net.Mail;

namespace RDLC_Demo.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;

        public HomeController(ILogger<HomeController> logger)
        {
            _logger = logger;
        }

        public IActionResult Index()
        {
            return View();
        }

        public IActionResult Privacy()
        {
            return View();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }

        public IActionResult ExportPDF()
        {
            var items = new[] {
                new Product { ProductID = 1, ProductName = "Adjustable Race", ProductNumber = "AR-5381" },
                new Product { ProductID = 2, ProductName = "Bearing Ball", ProductNumber = "BA-8327" }
            };

            var parameters = new[] { new ReportParameter("Title", "Hello ReportViewCore") };
            var reportFileName = "MyReports.Reports.Report1.rdlc";
            var dataSource = "DataSet1";
            var result = getReportFile(reportFileName, parameters, dataSource, items, "PDF");

            return File(result, "application/pdf", "Export.pdf");
        }

        public IActionResult ExportExcel()
        {
            var items = new[] {
                new Product { ProductID = 1, ProductName = "Adjustable Race", ProductNumber = "AR-5381" },
                new Product { ProductID = 2, ProductName = "Bearing Ball", ProductNumber = "BA-8327" }
            };

            var parameters = new[] { new ReportParameter("Title", "Hello ReportViewCore") };
            var reportFileName = "MyReports.Reports.Report1.rdlc";
            var dataSource = "DataSet1";
            var result = getReportFile(reportFileName, parameters, dataSource, items, "EXCEL");

            return File(result, "application/msexcel", "Export.xls");
        }

        public IActionResult SendMail()
        {
            var items = new[] {
                new Product { ProductID = 1, ProductName = "Adjustable Race", ProductNumber = "AR-5381" },
                new Product { ProductID = 2, ProductName = "Bearing Ball", ProductNumber = "BA-8327" }
            };

            var parameters = new[] { new ReportParameter("Title", "Hello ReportViewCore") };
            var reportFileName = "MyReports.Reports.Report1.rdlc";
            var dataSource = "DataSet1";
            var result = getReportFile(reportFileName, parameters, dataSource, items, "PDF");


            var emails = new[] { "rexmax1018@gmail.com" };

            doSendMail(emails, "Report Mail Test", "Test Content", result, "Export.pdf");

            return Content("Mail Sended.");
        }

        private byte[] getReportFile(string reportFileName, ReportParameter[] parameters, string dataSource, Array items, string fileType)
        {
            var report = new LocalReport();
            var assembly = typeof(Product).Assembly;

            using (var rs = assembly.GetManifestResourceStream(reportFileName))
            {
                report.LoadReportDefinition(rs);
                report.DataSources.Add(new ReportDataSource(dataSource, items));
                report.SetParameters(parameters);

                var result = report.Render(fileType);

                return result;
            }
        }

        private void doSendMail(string[] emails, string title, string content, byte[] fileBytes, string fileName)
        {
            var mail = new MailMessage();

            foreach (var email in emails)
            {
                mail.To.Add(email);
            }

            mail.Subject = title;
            mail.Body = content;
            mail.IsBodyHtml = true;
            mail.Priority = MailPriority.Normal;
            mail.From = new MailAddress("rexmax1018@gmail.com", "rexmax1018");

            if(fileBytes != null )
            {
                var att = new Attachment(new MemoryStream(fileBytes), fileName);
                mail.Attachments.Add(att);
            }

            var smtp = new SmtpClient("smtp.gmail.com", 587)
            {
                Credentials = new System.Net.NetworkCredential("rexmax1018@gmail.com", "password"),
                EnableSsl = true
            };

            smtp.Send(mail);

            mail.Dispose();
        }
    }
}
