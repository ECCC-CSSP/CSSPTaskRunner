using CSSPEnums;
using CSSPModels;
using CSSPTaskRunnerHelper.Resources;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Net.Mail;
using System.Security.Principal;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace CSSPTaskRunnerHelper
{
    public partial class TaskRunnerHelper
    {
        #region Variables
        #endregion Variables

        #region Properties
        #endregion Properties

        #region Constructors
        #endregion Constructors

        #region Functions public
        #endregion Functions public

        #region Functions private
        private void SendNewLabSheetEmailBigMPN(string href, TVItem tvItemProvince, TVItem tvItemSubsector, LabSheetAndA1Sheet labSheetAndA1Sheet)
        {
            using (CSSPWebToolsDBContext db = new CSSPWebToolsDBContext(DatabaseTypeEnum.SqlServerCSSPWebToolsDB))
            {
                SamplingPlan samplingPlan = (from c in db.SamplingPlans
                                             where c.SamplingPlanID == labSheetAndA1Sheet.LabSheet.SamplingPlanID
                                             select c).FirstOrDefault();

                List<SamplingPlanEmail> SamplingPlanEmailList = (from c in db.SamplingPlanEmails
                                                                 where c.SamplingPlanID == samplingPlan.SamplingPlanID
                                                                 orderby c.Email
                                                                 select c).ToList();

                if (!samplingPlan.IsActive)
                {
                    return;
                }

                // sending email to Contractors and Non Contractors
                foreach (bool IsContractor in new List<bool> { false, true })
                {
                    MailMessage mail = new MailMessage();

                    foreach (SamplingPlanEmail samplingPlanEmail in SamplingPlanEmailList.Where(c => c.IsContractor == IsContractor && c.LabSheetHasValueOver500 == true))
                    {
                        mail.To.Add(samplingPlanEmail.Email.ToLower());
                    }

                    if (mail.To.Count == 0)
                    {
                        continue;
                    }

                    mail.From = new MailAddress("ec.pccsm-cssp.ec@canada.ca");
                    mail.IsBodyHtml = true;

                    SmtpClient myClient = new System.Net.Mail.SmtpClient();

                    myClient.Host = "smtp.email-courriel.canada.ca";
                    myClient.Port = 587;
                    myClient.Credentials = new System.Net.NetworkCredential("ec.pccsm-cssp.ec@canada.ca", "H^9h6g@Gy$N57k=Dr@J7=F2y6p6b!T");
                    myClient.EnableSsl = true;

                    mail.Priority = MailPriority.High;

                    string TVTextSubsector = (from c in db.TVItemLanguages
                                              where c.TVItemID == tvItemSubsector.TVItemID
                                              && c.Language == LanguageEnum.en
                                              select c.TVText).FirstOrDefault();

                    if (string.IsNullOrWhiteSpace(TVTextSubsector))
                    {
                        TVTextSubsector = "TVTextSubsectorError";
                    }

                    int FirstSpace = TVTextSubsector.IndexOf(" ");

                    string subject = $"{ TVTextSubsector.Substring(0, (FirstSpace > 0 ? FirstSpace : TVTextSubsector.Length)) } – Lab sheet received – High MPN / Feuille de laboratoire reçu -  NPP élevé";

                    DateTime RunDate = new DateTime(int.Parse(labSheetAndA1Sheet.LabSheetA1Sheet.RunYear), int.Parse(labSheetAndA1Sheet.LabSheetA1Sheet.RunMonth), int.Parse(labSheetAndA1Sheet.LabSheetA1Sheet.RunDay));

                    StringBuilder msg = new StringBuilder();

                    // ---------------------- English part --------------

                    msg.AppendLine(@"<p>(français suit)</p>");
                    msg.AppendLine(@"<h2>Lab Sheet received</h2>");
                    msg.AppendLine($@"<h4>Subsector: { TVTextSubsector }</h4>");

                    Thread.CurrentThread.CurrentCulture = new CultureInfo("en-CA");
                    Thread.CurrentThread.CurrentUICulture = new CultureInfo("en-CA");

                    msg.AppendLine($@"<h4>Run Date: { RunDate.ToString("MMMM dd, yyyy") }</h4>");

                    if (!IsContractor)
                    {
                        msg.AppendLine($@"<a href=""{ href }"">Open CSSPWebTools</a>");
                    }
                    msg.AppendLine(@"<br /><br />");
                    msg.AppendLine(@"<p><b>Note: </b> Sites below are over the MPN threshold of 500</p>");

                    List<LabSheetA1Measurement> labSheetA1MeasurementList = (from c in labSheetAndA1Sheet.LabSheetA1Sheet.LabSheetA1MeasurementList
                                                                             where c.MPN != null
                                                                             && c.MPN >= MPNLimitForEmail
                                                                             select c).ToList();

                    msg.AppendLine("<ol>");
                    foreach (LabSheetA1Measurement labSheetA1Measurment in labSheetA1MeasurementList)
                    {
                        msg.AppendLine("<li>");
                        msg.AppendLine($"<b>Site: </b>{ labSheetA1Measurment.Site }{ (labSheetA1Measurment.SampleType == SampleTypeEnum.DailyDuplicate ? " Daily duplicate" : "") } <b>MPN: </b>{ labSheetA1Measurment.MPN }");
                        msg.AppendLine("</li>");

                    }
                    msg.AppendLine("</ol>");

                    msg.AppendLine(@"<br>");
                    msg.AppendLine(@"<p>Auto email from CSSPWebTools</p>");
                    msg.AppendLine(@"<br>");
                    msg.AppendLine(@"<hr />");

                    // ---------------------- French part --------------

                    msg.AppendLine(@"<hr />");
                    msg.AppendLine(@"<br>");
                    msg.AppendLine(@"<h2>Feuille de laboratoire reçu</h2>");
                    msg.AppendLine($@"<h4>Sous-secteur: { TVTextSubsector }</h4>");

                    Thread.CurrentThread.CurrentCulture = new CultureInfo("fr-CA");
                    Thread.CurrentThread.CurrentUICulture = new CultureInfo("fr-CA");

                    msg.AppendLine($@"<h4>Date de la tournée: { RunDate.ToString("dd MMMM, yyyy") }</h4>");

                    Thread.CurrentThread.CurrentCulture = new CultureInfo("en-CA");
                    Thread.CurrentThread.CurrentUICulture = new CultureInfo("en-CA");

                    if (!IsContractor)
                    {
                        msg.AppendLine($@"<a href=""{ href.Replace("en-CA", "fr-CA") }"">Open CSSPWebTools</a>");
                    }
                    msg.AppendLine(@"<br /><br />");
                    msg.AppendLine(@"<p><b>Remarque: </b> Les sites ci-dessous ont une valeure de NPP dépassant 500</p>");

                    msg.AppendLine("<ol>");
                    foreach (LabSheetA1Measurement labSheetA1Measurment in labSheetA1MeasurementList)
                    {
                        msg.AppendLine("<li>");
                        msg.AppendLine($"<b>Site: </b>{ labSheetA1Measurment.Site }{ (labSheetA1Measurment.SampleType == SampleTypeEnum.DailyDuplicate ? " Duplicata journalier" : "") } <b>NPP: </b>{ labSheetA1Measurment.MPN }");
                        msg.AppendLine("</li>");

                    }
                    msg.AppendLine("</ol>");

                    msg.AppendLine(@"<br>");
                    msg.AppendLine(@"<p>Courriel automatique provenant de CSSPWebTools</p>");

                    mail.Subject = subject;
                    mail.Body = msg.ToString();
                    myClient.Send(mail);
                }
            }
        }
        public void SendNewLabSheetEmail(string href, TVItem tvItemProvince, TVItem tvItemSubsector, LabSheetAndA1Sheet labSheetAndA1Sheet)
        {
            using (CSSPWebToolsDBContext db = new CSSPWebToolsDBContext(DatabaseTypeEnum.SqlServerCSSPWebToolsDB))
            {
                SamplingPlan samplingPlan = (from c in db.SamplingPlans
                                             where c.SamplingPlanID == labSheetAndA1Sheet.LabSheet.SamplingPlanID
                                             select c).FirstOrDefault();

                List<SamplingPlanEmail> SamplingPlanEmailList = (from c in db.SamplingPlanEmails
                                                                 where c.SamplingPlanID == samplingPlan.SamplingPlanID
                                                                 orderby c.Email
                                                                 select c).ToList();

                if (!samplingPlan.IsActive)
                {
                    return;
                }


                // sending email to Non Contractors

                foreach (bool IsContractor in new List<bool> { false, true })
                {
                    MailMessage mail = new MailMessage();

                    foreach (SamplingPlanEmail samplingPlanEmail in SamplingPlanEmailList.Where(c => c.IsContractor == IsContractor && c.LabSheetReceived == true))
                    {
                        mail.To.Add(samplingPlanEmail.Email.ToLower());
                    }

                    if (mail.To.Count == 0)
                    {
                        continue;
                    }

                    mail.From = new MailAddress("ec.pccsm-cssp.ec@canada.ca");
                    mail.IsBodyHtml = true;

                    SmtpClient myClient = new System.Net.Mail.SmtpClient();

                    myClient.Host = "smtp.email-courriel.canada.ca";
                    myClient.Port = 587;
                    myClient.Credentials = new System.Net.NetworkCredential("ec.pccsm-cssp.ec@canada.ca", "H^9h6g@Gy$N57k=Dr@J7=F2y6p6b!T");
                    myClient.EnableSsl = true;

                    string TVTextSubsector = (from c in db.TVItemLanguages
                                              where c.TVItemID == tvItemSubsector.TVItemID
                                              && c.Language == LanguageEnum.en
                                              select c.TVText).FirstOrDefault();

                    if (string.IsNullOrWhiteSpace(TVTextSubsector))
                    {
                        TVTextSubsector = "TVTextSubsectorError";
                    }

                    int FirstSpace = TVTextSubsector.IndexOf(" ");

                    string subject = $"{ TVTextSubsector.Substring(0, (FirstSpace > 0 ? FirstSpace : TVTextSubsector.Length)) } – Lab sheet received / Feuille de laboratoire reçu";

                    DateTime RunDate = new DateTime(int.Parse(labSheetAndA1Sheet.LabSheetA1Sheet.RunYear), int.Parse(labSheetAndA1Sheet.LabSheetA1Sheet.RunMonth), int.Parse(labSheetAndA1Sheet.LabSheetA1Sheet.RunDay));

                    StringBuilder msg = new StringBuilder();

                    // ---------------------- English part --------------

                    msg.AppendLine(@"<p>(français suit)</p>");
                    msg.AppendLine(@"<h2>Lab Sheet received</h2>");
                    msg.AppendLine($@"<h4>Subsector: { TVTextSubsector }</h4>");
                    msg.AppendLine($@"<h4>Run Date: { RunDate.ToString("MMMM dd, yyyy") }</h4>");
                    if (!IsContractor)
                    {
                        msg.AppendLine($@"<a href=""{ href }"">Open CSSPWebTools</a>");
                    }

                    msg.AppendLine(@"<br>");
                    msg.AppendLine(@"<p>Auto email from CSSPWebTools</p>");
                    msg.AppendLine(@"<br>");
                    msg.AppendLine(@"<hr />");

                    // ---------------------- French part --------------

                    msg.AppendLine(@"<hr />");
                    msg.AppendLine(@"<br>");
                    msg.AppendLine(@"<h2>Feuille de laboratoire reçu</h2>");
                    msg.AppendLine($@"<h4>Sous-secteur: { TVTextSubsector }</h4>");

                    Thread.CurrentThread.CurrentCulture = new CultureInfo("fr-CA");
                    Thread.CurrentThread.CurrentUICulture = new CultureInfo("fr-CA");

                    msg.AppendLine($@"<h4>Date de la tournée: { RunDate.ToString("dd MMMM, yyyy") }</h4>");

                    Thread.CurrentThread.CurrentCulture = new CultureInfo("en-CA");
                    Thread.CurrentThread.CurrentUICulture = new CultureInfo("en-CA");

                    if (!IsContractor)
                    {
                        msg.AppendLine($@"<a href=""{ href.Replace("en-CA", "fr-CA") }"">Open CSSPWebTools</a>");
                    }

                    msg.AppendLine(@"<br>");
                    msg.AppendLine(@"<p>Courriel automatique provenant de CSSPWebTools</p>");

                    mail.Subject = subject;
                    mail.Body = msg.ToString();
                    myClient.Send(mail);
                }
            }
        }
        private void ShawnDonohueEmail(Exception ex)
        {
            try
            {
                MailMessage mail = new MailMessage();

                mail.To.Add("Shawn.Donohue@Canada.ca");
                mail.Bcc.Add("Charles.LeBlanc2@Canada.ca");

                mail.From = new MailAddress("ec.pccsm-cssp.ec@canada.ca");
                mail.IsBodyHtml = true;

                SmtpClient myClient = new System.Net.Mail.SmtpClient();

                myClient.Host = "smtp.email-courriel.canada.ca";
                myClient.Port = 587;
                myClient.Credentials = new System.Net.NetworkCredential("ec.pccsm-cssp.ec@canada.ca", "H^9h6g@Gy$N57k=Dr@J7=F2y6p6b!T");
                myClient.EnableSsl = true;

                string subject = "Shawn Donohue Issue from CSSPWebToolsTaskRunner";

                StringBuilder msg = new StringBuilder();

                msg.AppendLine("<h2>Shawn Donohue Issue Email</h2>");
                msg.AppendLine($"<h4>Date of issue: { DateTime.Now }</h4>");
                msg.AppendLine($"<h4>Exception Message: { ex.Message }</h4>");
                msg.AppendLine($"<h4>Exception Inner Message: { (ex.InnerException != null ? ex.InnerException.Message : "empty") }</h4>");

                msg.AppendLine(@"<br>");
                msg.AppendLine(@"<p>Auto email from CSSPWebTools.</p>");

                mail.Subject = subject;
                mail.Body = msg.ToString();
                myClient.Send(mail);
            }
            catch (Exception ex2)
            {
                string InnerException = (ex2.InnerException == null ? "" : $" InnerException: { ex2.InnerException.Message }");
                MessageEvent(new MessageEventArgs($"Could not send ShawDonohueEmail. Error: { ex2.Message }{ InnerException }", MessageTypeEnum.Error));
            }
        }
        #endregion Functions private
    }
}
