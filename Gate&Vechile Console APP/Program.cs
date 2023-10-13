using System;
using System.Globalization;
using System.IO;
using System.Net;
using System.Net.Mail;
using System.Xml.Linq;
using OfficeOpenXml;
using Serilog.Events;
using Serilog;

namespace GatePassExpiryNotifier
{
    class Program
    {
        static void Main(string[] args)
        {
           //Enter your Excel Path Here
            string excelFilePath = "";
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            if (File.Exists(excelFilePath))
            {
                ProcessGatePassSheet(excelFilePath);
                ProcessVehiclePassSheet(excelFilePath);
            }
            else
            {
                Console.WriteLine("Excel file not found at the specified path.");
            }
        }

        

        static void SendGatePassEmail(string name, string toEmail, string gatePass, DateTime toDate, DateTime pvcDate)
        {
            // Configure your SMTP settings
            string smtpServer = "";
            int smtpPort ;
            string smtpUsername = "";
            string smtpPassword = "";

            MailMessage mailMessage = new MailMessage();

            

            using (SmtpClient smtpClient = new SmtpClient(smtpServer, smtpPort))
            {
                smtpClient.UseDefaultCredentials = false;
                smtpClient.Credentials = new NetworkCredential(smtpUsername, smtpPassword);
                smtpClient.EnableSsl = true;

                string subject = string.Empty;
                string body = $"Dear {name},<br/><br/>";
                try
                {
                    if (DateTime.Now > toDate && DateTime.Now > pvcDate)
                    {
                        // Both pass types have already expired
                        body += $"GatePass & PVC has expired. Gate pass no:{gatePass}.";
                        body += $"Expiry Date: {toDate.ToString("dd-MMM-yy")}<br/>";
                        subject = "Expired Gate Pass and PVC Notification";
                    }
                    else if (DateTime.Now > toDate && (pvcDate - DateTime.Now).TotalDays <= 15)
                    {
                        // Only gate pass has expired
                        subject = "Gate Pass Expired and Upcoming PVC expiry Notification";
                        body += $"GatePass is expired for {name}. GatePass NO: {gatePass}.<br>";
                        body += $"Expiry Date: {toDate.ToString("dd-MMM-yy")}<br>";
                        body += $"PVC is expiring soon. PVC Expiry Date: {pvcDate.ToString("dd-MMM-yy")}<br/>";
                    }
                    else if (DateTime.Now > pvcDate && (toDate - DateTime.Now).TotalDays <= 7)
                    {
                        // Only PVC has expired
                        subject = "Expired PVC and Upcoming GatePass expiry Notification";
                        body += $"The PVC is expired on {pvcDate.ToString("dd-MMM-yy")}.<br> Gate Pass NO: {gatePass} is expiring soon.<br/>";

                    }
                    else if (DateTime.Now > toDate && DateTime.Now < pvcDate)
                    {
                        subject = "Expired GatePass Notification";
                        body += $"GatePass is Expired , Gate Pass No: {gatePass}.<br/>";
                        body += $"Expiry Date: {toDate.ToString("dd-MMM-yy")}<br/>";

                    }
                    else if (DateTime.Now > pvcDate && DateTime.Now < toDate)
                    {
                        subject = "Expired PVC Notification";
                        body += $"PVC is Expired. PVC Expiry Date: {pvcDate.ToString("dd-MMM-yy")}<br/>";
                    }
                    else if ((toDate - DateTime.Now).TotalDays <= 7 && (pvcDate - DateTime.Now).TotalDays <= 15)
                    {
                        subject = "Gate Pass and PVC Expiry Notification";
                        //body += $"The gate pass and PVC are expiring soon. Gate Pass No: {gatePass}, PVC Expiry Date: {pvcDate.ToString("dd-MMM-yy")}.<br/>";
                        body += $"Your gate pass  is expiring soon. Gate Pass No: {gatePass}.<br/>";
                        body += $"Your PVC is expiring soon. PVC Expiry Date: {pvcDate.ToString("dd-MMM-yy", CultureInfo.InvariantCulture)}.<br/>";
                    

                }
                    else if ((toDate - DateTime.Now).TotalDays <= 7)
                    {
                        // Only gate pass will expire within 7 days
                        subject = "Gate Pass Expiry Notification";
                        body += $"The gate pass for {name} is expiring soon. Gate Pass No: {gatePass}.<br/>";
                        body += $"Expiry Date: {toDate.ToString("dd-MMM-yy")}<br/>";
                    }
                    else
                    {
                        // Only PVC will expire within 15 days
                        subject = "PVC Expiry Notification";
                        body += $"The PVC for {name} is expiring soon.<br/>"; 
                        body += $"PVC Expiry Date: {pvcDate.ToString("dd-MMM-yy")}<br/>";
                    }
                }
                catch (Exception ex)
                {
                    string err = ex.Message;
                }


                body += "<br/>Thanks & Regards,<br/>";
                body += "";

                mailMessage.Body = body;
                mailMessage.Subject = subject;

                mailMessage.From = new MailAddress("J");
                mailMessage.To.Add(toEmail);
                mailMessage.CC.Add("");
                mailMessage.IsBodyHtml = true;

                try
                {
                    smtpClient.Send(mailMessage);
                    Console.WriteLine($"Pass expiry email sent to {name} ({toEmail}).");
                    Console.Clear();
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to send pass expiry email to {name} ({toEmail})");
                }
            }
        }




        static void SendVehiclePassEmail(string name, string toEmail, string vehicleNo, DateTime toDate)
        {
            // Configure your SMTP settings
            string smtpServer = "";
            int smtpPort ;
            string smtpUsername = "";
            string smtpPassword = "";

            using (SmtpClient smtpClient = new SmtpClient(smtpServer, smtpPort))
            {
                smtpClient.UseDefaultCredentials = false;
                smtpClient.Credentials = new NetworkCredential(smtpUsername, smtpPassword);
                smtpClient.EnableSsl = true;

                // Create the email message
                MailMessage mailMessage = new MailMessage();
                mailMessage.From = new MailAddress("");
                mailMessage.To.Add(toEmail);
                mailMessage.CC.Add("");
                mailMessage.IsBodyHtml = true;

                if (DateTime.Now > toDate)
                {
                    // Pass has already expired
                    mailMessage.Subject = "Expired Vehicle Pass Notification";

                    string body = $"Dear {name},<br/><br/>";
                    body += $"The vehicle pass for Vehicle No: {vehicleNo} has already expired on {toDate:dd-MMM-yy}.<br/>";
                    body += "Please renew your vehicle pass.<br/><br/>";
                    body += "Thanks & Regards,<br/>";
                    body += "Kundan Parkar";

                    mailMessage.Body = body;
                }
                else if ((toDate - DateTime.Now).TotalDays <= 7)
                {
                    // Pass will expire within 7 days
                    mailMessage.Subject = "Vehicle Pass Expiry Notification";

                    string body = $"Dear Sir/Madam,<br/><br/>";
                    body += $"The vehicle pass for {name} is expiring soon. Vehicle No: {vehicleNo}.<br/>";
                    body += $"Expiry Date: {toDate.ToString("dd-MMM-yy")}<br/><br/>";
                    body += "Thanks & Regards,<br/>";
                    body += "";

                    mailMessage.Body = body;
                }

                // Send the email
                try
                {
                    smtpClient.Send(mailMessage);
                    Console.WriteLine($"Vehicle Pass email sent to {name} ({toEmail}).");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to send Vehicle Pass email to {name} ({toEmail}): {ex.Message}");
                }
            }
        }

        static void ProcessGatePassSheet(string excelFilePath)
        {
            FileInfo fileInfo = new FileInfo(excelFilePath);
            try
            {
                using (ExcelPackage package = new ExcelPackage(fileInfo))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets["gatepass"]; 
                    int toColumnIndex = 6; 
                    DateTime currentDate = DateTime.Now;

                    for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
                    {
                        string dateText = worksheet.Cells[row, toColumnIndex].Text;
                        string gatePass = worksheet.Cells[row, 2].Text;
                        string email = worksheet.Cells[row, 9].Text; 
                        string name = worksheet.Cells[row, 3].Text; 

                        if (DateTime.TryParseExact(dateText, "d-MMM-yy", CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime toDate))
                        {
                            DateTime pvcDate; // Assuming PVC date is in column 7 (index 6)
                            string pvcDateText = worksheet.Cells[row, 7].Text;

                            if (DateTime.TryParseExact(pvcDateText, "d-MMM-yy", CultureInfo.InvariantCulture, DateTimeStyles.None, out pvcDate))
                            {
                                //if ((toDate - currentDate).TotalDays <= 7 || (pvcDate - currentDate).TotalDays <= 15)
                                //{
                                // At least one pass type will expire within 7 days(for gate pass) or 15 days(for PVC)


                                if (DateTime.Now > toDate && DateTime.Now > pvcDate)
                                {
                                    SendGatePassEmail(name, email, gatePass, toDate, pvcDate);
                                }
                                else if ((toDate - currentDate).TotalDays <= 7)
                                {
                                    SendGatePassEmail(name, email, gatePass, toDate, pvcDate);
                                }
                                else if ((pvcDate - currentDate).TotalDays <= 15)
                                {
                                    SendGatePassEmail(name, email, gatePass, toDate, pvcDate);
                                }
                                else if (DateTime.Now > toDate)
                                {
                                    SendGatePassEmail(name, email, gatePass, toDate, pvcDate);
                                }
                                else if (DateTime.Now > pvcDate)
                                {
                                    SendGatePassEmail(name, email, gatePass, toDate, pvcDate);
                                }

                                else
                                {
                                   // Console.WriteLine($"Invalid date format at row {row}, column {toColumnIndex}. Value: {dateText}");
                                }
                            }


                            else
                            {
                                //Console.WriteLine($"Invalid PVC date format at row {row}, column 7. Value: {pvcDateText}");
                            }
                        }
                        else
                        {
                           // Console.WriteLine($"could send mail for {name}");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing Gate Pass sheet");
            }
        }



        static void ProcessVehiclePassSheet(string excelFilePath)
        {
            FileInfo fileInfo = new FileInfo(excelFilePath);
            try
            {
                using (ExcelPackage package = new ExcelPackage(fileInfo))
                {
;                    ExcelWorksheet worksheet = package.Workbook.Worksheets["vechilepass"]; // Assuming sheet name is "vechilepass"

                    int toColumnIndex = 5; // Assuming the "To" column is in column 6 (index 5)
                    DateTime currentDate = DateTime.Now;

                    for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
                    {
                        string dateText = worksheet.Cells[row, toColumnIndex].Text;
                        if (DateTime.TryParseExact(dateText, "d-MMM-yy", CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime toDate))
                        {
                            if ((toDate - currentDate).TotalDays <= 7)
                            {
                                string vehicleNo = worksheet.Cells[row, 2].Text; 
                                string email = worksheet.Cells[row, 8].Text; 
                                string name = worksheet.Cells[row, 3].Text; 

                                SendVehiclePassEmail(name, email, vehicleNo, toDate);
                            }
                        }
                        else if ((toDate - currentDate).TotalDays <= 7)
                        {
                            // Pass will expire within 7 days
                            string vehicleNo = worksheet.Cells[row, 2].Text;
                            string email = worksheet.Cells[row, 8].Text; 
                            string name = worksheet.Cells[row, 3].Text; 

                            // Send notification for pass about to expire
                            if (email != "")
                            {
                                SendVehiclePassEmail(name, email, vehicleNo, toDate);
                            }
                            else
                            {
                                Console.WriteLine("Email not found for " + name);
                            }
                        }
                        else
                        {
                            Console.WriteLine($"Invalid date format at row {row}, column {toColumnIndex}. Value: {dateText}");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                string err = ex.Message;
                Console.WriteLine("Error processing Vehicle Pass sheet");
            }
        }
    }
}
