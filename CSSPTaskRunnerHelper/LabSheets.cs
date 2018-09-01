using CSSPEnums;
using CSSPModels;
using CSSPServices;
using CSSPTaskRunnerHelper.Resources;
using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Linq;
using System.Net;
using System.Security.Principal;
using System.Text;
using System.Threading.Tasks;
using System.Transactions;

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
        public string UpdateOtherServerWithOtherServerLabSheetIDAndLabSheetStatus(int OtherServerLabSheetID, LabSheetStatusEnum LabSheetStatus)
        {
            string retStr = "";

            NameValueCollection paramList = new NameValueCollection();
            paramList.Add("OtherServerLabSheetID", OtherServerLabSheetID.ToString());
            paramList.Add("LabSheetStatus", ((int)LabSheetStatus).ToString());

            using (WebClient webClient = new WebClient())
            {
                WebProxy webProxy = new WebProxy();
                webClient.Proxy = webProxy;

                webClient.Headers.Add("Content-Type", "application/x-www-form-urlencoded");
                byte[] ret = webClient.UploadValues(new Uri("http://cssplabsheet.azurewebsites.net/UpdateLabSheetStatus.aspx"), "POST", paramList);

                retStr = System.Text.Encoding.Default.GetString(ret);
            }

            return retStr;
        }
        public string UploadLabSheetDetailInDB(LabSheetAndA1Sheet labSheetAndA1Sheet)
        {
            string retStr = "";

            // Filling LabSheetDetailModel
            LabSheetDetail labSheetDetailNew = new LabSheetDetail();

            labSheetDetailNew.LabSheetID = labSheetAndA1Sheet.LabSheet.LabSheetID;
            labSheetDetailNew.SamplingPlanID = labSheetAndA1Sheet.LabSheet.SamplingPlanID;
            labSheetDetailNew.SubsectorTVItemID = labSheetAndA1Sheet.LabSheet.SubsectorTVItemID;
            labSheetDetailNew.Version = labSheetAndA1Sheet.LabSheetA1Sheet.Version;
            labSheetDetailNew.IncludeLaboratoryQAQC = labSheetAndA1Sheet.LabSheetA1Sheet.IncludeLaboratoryQAQC;

            // RunDate
            DateTime RunDate = new DateTime();
            int RunYear = (string.IsNullOrWhiteSpace(labSheetAndA1Sheet.LabSheetA1Sheet.RunYear) ? 1900 : int.Parse(labSheetAndA1Sheet.LabSheetA1Sheet.RunYear));
            int RunMonth = (string.IsNullOrWhiteSpace(labSheetAndA1Sheet.LabSheetA1Sheet.RunMonth) ? 1 : int.Parse(labSheetAndA1Sheet.LabSheetA1Sheet.RunMonth));
            int RunDay = (string.IsNullOrWhiteSpace(labSheetAndA1Sheet.LabSheetA1Sheet.RunDay) ? 1 : int.Parse(labSheetAndA1Sheet.LabSheetA1Sheet.RunDay));
            if (RunYear != 1900)
            {
                RunDate = new DateTime(RunYear, RunMonth, RunDay);
            }

            labSheetDetailNew.RunDate = RunDate;

            labSheetDetailNew.Tides = labSheetAndA1Sheet.LabSheetA1Sheet.Tides;
            labSheetDetailNew.SampleCrewInitials = labSheetAndA1Sheet.LabSheetA1Sheet.SampleCrewInitials;

            if (labSheetAndA1Sheet.LabSheetA1Sheet.IncludeLaboratoryQAQC)
            {
                // IncubationBath1StartDate
                DateTime IncubationBath1StartDate = new DateTime();
                int IncubationBath1StartYear = (string.IsNullOrWhiteSpace(labSheetAndA1Sheet.LabSheetA1Sheet.RunYear) ? 1900 : int.Parse(labSheetAndA1Sheet.LabSheetA1Sheet.RunYear));
                int IncubationBath1StartMonth = (string.IsNullOrWhiteSpace(labSheetAndA1Sheet.LabSheetA1Sheet.RunMonth) ? 1 : int.Parse(labSheetAndA1Sheet.LabSheetA1Sheet.RunMonth));
                int IncubationBath1StartDay = (string.IsNullOrWhiteSpace(labSheetAndA1Sheet.LabSheetA1Sheet.RunDay) ? 1 : int.Parse(labSheetAndA1Sheet.LabSheetA1Sheet.RunDay));
                int IncubationBath1StartHour = 0;
                int IncubationBath1StartMinute = 0;
                if (labSheetAndA1Sheet.LabSheetA1Sheet.IncubationBath1StartTime.Length == 5)
                {
                    IncubationBath1StartHour = int.Parse(labSheetAndA1Sheet.LabSheetA1Sheet.IncubationBath1StartTime.Substring(0, 2));
                    IncubationBath1StartMinute = int.Parse(labSheetAndA1Sheet.LabSheetA1Sheet.IncubationBath1StartTime.Substring(3, 2));
                }
                if (IncubationBath1StartYear != 1900)
                {
                    IncubationBath1StartDate = new DateTime(IncubationBath1StartYear, IncubationBath1StartMonth, IncubationBath1StartDay, IncubationBath1StartHour, IncubationBath1StartMinute, 0);
                }

                labSheetDetailNew.IncubationBath1StartTime = IncubationBath1StartDate;

                // IncubationBath1EndDate
                DateTime IncubationBath1EndDate = new DateTime();
                int IncubationBath1EndYear = (string.IsNullOrWhiteSpace(labSheetAndA1Sheet.LabSheetA1Sheet.RunYear) ? 1900 : int.Parse(labSheetAndA1Sheet.LabSheetA1Sheet.RunYear));
                int IncubationBath1EndMonth = (string.IsNullOrWhiteSpace(labSheetAndA1Sheet.LabSheetA1Sheet.RunMonth) ? 1 : int.Parse(labSheetAndA1Sheet.LabSheetA1Sheet.RunMonth));
                int IncubationBath1EndDay = (string.IsNullOrWhiteSpace(labSheetAndA1Sheet.LabSheetA1Sheet.RunDay) ? 1 : int.Parse(labSheetAndA1Sheet.LabSheetA1Sheet.RunDay));
                int IncubationBath1EndHour = 0;
                int IncubationBath1EndMinute = 0;
                if (labSheetAndA1Sheet.LabSheetA1Sheet.IncubationBath1EndTime.Length == 5)
                {
                    IncubationBath1EndHour = int.Parse(labSheetAndA1Sheet.LabSheetA1Sheet.IncubationBath1EndTime.Substring(0, 2));
                    IncubationBath1EndMinute = int.Parse(labSheetAndA1Sheet.LabSheetA1Sheet.IncubationBath1EndTime.Substring(3, 2));
                }
                if (IncubationBath1EndYear != 1900)
                {
                    IncubationBath1EndDate = new DateTime(IncubationBath1EndYear, IncubationBath1EndMonth, IncubationBath1EndDay, IncubationBath1EndHour, IncubationBath1EndMinute, 0).AddDays(1);
                }

                labSheetDetailNew.IncubationBath1EndTime = IncubationBath1EndDate;

                // IncubationBath1TimeCalculated_minutes 
                int IncubationBath1TimeCalculated_minutes = 0;
                if (labSheetAndA1Sheet.LabSheetA1Sheet.IncubationBath1TimeCalculated.Length == 5)
                {
                    IncubationBath1TimeCalculated_minutes = int.Parse(labSheetAndA1Sheet.LabSheetA1Sheet.IncubationBath1TimeCalculated.Substring(0, 2)) * 60;
                    IncubationBath1TimeCalculated_minutes += int.Parse(labSheetAndA1Sheet.LabSheetA1Sheet.IncubationBath1TimeCalculated.Substring(3, 2));
                }

                labSheetDetailNew.IncubationBath1TimeCalculated_minutes = IncubationBath1TimeCalculated_minutes;

                labSheetDetailNew.WaterBath1 = labSheetAndA1Sheet.LabSheetA1Sheet.WaterBath1;

                labSheetDetailNew.WaterBathCount = labSheetAndA1Sheet.LabSheetA1Sheet.WaterBathCount;

                if (labSheetDetailNew.WaterBathCount > 1)
                {
                    // IncubationBath2StartDate
                    DateTime IncubationBath2StartDate = new DateTime();
                    int IncubationBath2StartYear = (string.IsNullOrWhiteSpace(labSheetAndA1Sheet.LabSheetA1Sheet.RunYear) ? 1900 : int.Parse(labSheetAndA1Sheet.LabSheetA1Sheet.RunYear));
                    int IncubationBath2StartMonth = (string.IsNullOrWhiteSpace(labSheetAndA1Sheet.LabSheetA1Sheet.RunMonth) ? 1 : int.Parse(labSheetAndA1Sheet.LabSheetA1Sheet.RunMonth));
                    int IncubationBath2StartDay = (string.IsNullOrWhiteSpace(labSheetAndA1Sheet.LabSheetA1Sheet.RunDay) ? 1 : int.Parse(labSheetAndA1Sheet.LabSheetA1Sheet.RunDay));
                    int IncubationBath2StartHour = 0;
                    int IncubationBath2StartMinute = 0;
                    if (labSheetAndA1Sheet.LabSheetA1Sheet.IncubationBath2StartTime.Length == 5)
                    {
                        IncubationBath2StartHour = int.Parse(labSheetAndA1Sheet.LabSheetA1Sheet.IncubationBath2StartTime.Substring(0, 2));
                        IncubationBath2StartMinute = int.Parse(labSheetAndA1Sheet.LabSheetA1Sheet.IncubationBath2StartTime.Substring(3, 2));
                    }
                    if (IncubationBath2StartYear != 1900)
                    {
                        IncubationBath2StartDate = new DateTime(IncubationBath2StartYear, IncubationBath2StartMonth, IncubationBath2StartDay, IncubationBath2StartHour, IncubationBath2StartMinute, 0);
                    }

                    labSheetDetailNew.IncubationBath2StartTime = IncubationBath2StartDate;

                    // IncubationBath2EndDate
                    DateTime IncubationBath2EndDate = new DateTime();
                    int IncubationBath2EndYear = (string.IsNullOrWhiteSpace(labSheetAndA1Sheet.LabSheetA1Sheet.RunYear) ? 1900 : int.Parse(labSheetAndA1Sheet.LabSheetA1Sheet.RunYear));
                    int IncubationBath2EndMonth = (string.IsNullOrWhiteSpace(labSheetAndA1Sheet.LabSheetA1Sheet.RunMonth) ? 1 : int.Parse(labSheetAndA1Sheet.LabSheetA1Sheet.RunMonth));
                    int IncubationBath2EndDay = (string.IsNullOrWhiteSpace(labSheetAndA1Sheet.LabSheetA1Sheet.RunDay) ? 1 : int.Parse(labSheetAndA1Sheet.LabSheetA1Sheet.RunDay));
                    int IncubationBath2EndHour = 0;
                    int IncubationBath2EndMinute = 0;
                    if (labSheetAndA1Sheet.LabSheetA1Sheet.IncubationBath2EndTime.Length == 5)
                    {
                        IncubationBath2EndHour = int.Parse(labSheetAndA1Sheet.LabSheetA1Sheet.IncubationBath2EndTime.Substring(0, 2));
                        IncubationBath2EndMinute = int.Parse(labSheetAndA1Sheet.LabSheetA1Sheet.IncubationBath2EndTime.Substring(3, 2));
                    }
                    if (IncubationBath2EndYear != 1900)
                    {
                        IncubationBath2EndDate = new DateTime(IncubationBath2EndYear, IncubationBath2EndMonth, IncubationBath2EndDay, IncubationBath2EndHour, IncubationBath2EndMinute, 0).AddDays(1);
                    }

                    labSheetDetailNew.IncubationBath2EndTime = IncubationBath2EndDate;

                    // IncubationBath2TimeCalculated_minutes 
                    int IncubationBath2TimeCalculated_minutes = 0;
                    if (labSheetAndA1Sheet.LabSheetA1Sheet.IncubationBath2TimeCalculated.Length == 5)
                    {
                        IncubationBath2TimeCalculated_minutes = int.Parse(labSheetAndA1Sheet.LabSheetA1Sheet.IncubationBath2TimeCalculated.Substring(0, 2)) * 60;
                        IncubationBath2TimeCalculated_minutes += int.Parse(labSheetAndA1Sheet.LabSheetA1Sheet.IncubationBath2TimeCalculated.Substring(3, 2));
                    }

                    labSheetDetailNew.IncubationBath2TimeCalculated_minutes = IncubationBath2TimeCalculated_minutes;

                    labSheetDetailNew.WaterBath2 = labSheetAndA1Sheet.LabSheetA1Sheet.WaterBath2;
                    labSheetDetailNew.Bath2Positive44_5 = labSheetAndA1Sheet.LabSheetA1Sheet.Bath2Positive44_5;
                    labSheetDetailNew.Bath2NonTarget44_5 = labSheetAndA1Sheet.LabSheetA1Sheet.Bath2NonTarget44_5;
                    labSheetDetailNew.Bath2Negative44_5 = labSheetAndA1Sheet.LabSheetA1Sheet.Bath2Negative44_5;
                    labSheetDetailNew.Bath2Blank44_5 = labSheetAndA1Sheet.LabSheetA1Sheet.Bath2Blank44_5;

                }

                if (labSheetDetailNew.WaterBathCount > 2)
                {
                    // IncubationBath3StartDate
                    DateTime IncubationBath3StartDate = new DateTime();
                    int IncubationBath3StartYear = (string.IsNullOrWhiteSpace(labSheetAndA1Sheet.LabSheetA1Sheet.RunYear) ? 1900 : int.Parse(labSheetAndA1Sheet.LabSheetA1Sheet.RunYear));
                    int IncubationBath3StartMonth = (string.IsNullOrWhiteSpace(labSheetAndA1Sheet.LabSheetA1Sheet.RunMonth) ? 1 : int.Parse(labSheetAndA1Sheet.LabSheetA1Sheet.RunMonth));
                    int IncubationBath3StartDay = (string.IsNullOrWhiteSpace(labSheetAndA1Sheet.LabSheetA1Sheet.RunDay) ? 1 : int.Parse(labSheetAndA1Sheet.LabSheetA1Sheet.RunDay));
                    int IncubationBath3StartHour = 0;
                    int IncubationBath3StartMinute = 0;
                    if (labSheetAndA1Sheet.LabSheetA1Sheet.IncubationBath3StartTime.Length == 5)
                    {
                        IncubationBath3StartHour = int.Parse(labSheetAndA1Sheet.LabSheetA1Sheet.IncubationBath3StartTime.Substring(0, 2));
                        IncubationBath3StartMinute = int.Parse(labSheetAndA1Sheet.LabSheetA1Sheet.IncubationBath3StartTime.Substring(3, 2));
                    }
                    if (IncubationBath3StartYear != 1900)
                    {
                        IncubationBath3StartDate = new DateTime(IncubationBath3StartYear, IncubationBath3StartMonth, IncubationBath3StartDay, IncubationBath3StartHour, IncubationBath3StartMinute, 0);
                    }

                    labSheetDetailNew.IncubationBath3StartTime = IncubationBath3StartDate;

                    // IncubationBath3EndDate
                    DateTime IncubationBath3EndDate = new DateTime();
                    int IncubationBath3EndYear = (string.IsNullOrWhiteSpace(labSheetAndA1Sheet.LabSheetA1Sheet.RunYear) ? 1900 : int.Parse(labSheetAndA1Sheet.LabSheetA1Sheet.RunYear));
                    int IncubationBath3EndMonth = (string.IsNullOrWhiteSpace(labSheetAndA1Sheet.LabSheetA1Sheet.RunMonth) ? 1 : int.Parse(labSheetAndA1Sheet.LabSheetA1Sheet.RunMonth));
                    int IncubationBath3EndDay = (string.IsNullOrWhiteSpace(labSheetAndA1Sheet.LabSheetA1Sheet.RunDay) ? 1 : int.Parse(labSheetAndA1Sheet.LabSheetA1Sheet.RunDay));
                    int IncubationBath3EndHour = 0;
                    int IncubationBath3EndMinute = 0;
                    if (labSheetAndA1Sheet.LabSheetA1Sheet.IncubationBath3EndTime.Length == 5)
                    {
                        IncubationBath3EndHour = int.Parse(labSheetAndA1Sheet.LabSheetA1Sheet.IncubationBath3EndTime.Substring(0, 2));
                        IncubationBath3EndMinute = int.Parse(labSheetAndA1Sheet.LabSheetA1Sheet.IncubationBath3EndTime.Substring(3, 2));
                    }
                    if (IncubationBath3EndYear != 1900)
                    {
                        IncubationBath3EndDate = new DateTime(IncubationBath3EndYear, IncubationBath3EndMonth, IncubationBath3EndDay, IncubationBath3EndHour, IncubationBath3EndMinute, 0).AddDays(1);
                    }

                    labSheetDetailNew.IncubationBath3EndTime = IncubationBath3EndDate;

                    // IncubationBath3TimeCalculated_minutes 
                    int IncubationBath3TimeCalculated_minutes = 0;
                    if (labSheetAndA1Sheet.LabSheetA1Sheet.IncubationBath3TimeCalculated.Length == 5)
                    {
                        IncubationBath3TimeCalculated_minutes = int.Parse(labSheetAndA1Sheet.LabSheetA1Sheet.IncubationBath3TimeCalculated.Substring(0, 2)) * 60;
                        IncubationBath3TimeCalculated_minutes += int.Parse(labSheetAndA1Sheet.LabSheetA1Sheet.IncubationBath3TimeCalculated.Substring(3, 2));
                    }

                    labSheetDetailNew.IncubationBath3TimeCalculated_minutes = IncubationBath3TimeCalculated_minutes;

                    labSheetDetailNew.WaterBath3 = labSheetAndA1Sheet.LabSheetA1Sheet.WaterBath3;
                    labSheetDetailNew.Bath3Positive44_5 = labSheetAndA1Sheet.LabSheetA1Sheet.Bath3Positive44_5;
                    labSheetDetailNew.Bath3NonTarget44_5 = labSheetAndA1Sheet.LabSheetA1Sheet.Bath3NonTarget44_5;
                    labSheetDetailNew.Bath3Negative44_5 = labSheetAndA1Sheet.LabSheetA1Sheet.Bath3Negative44_5;
                    labSheetDetailNew.Bath3Blank44_5 = labSheetAndA1Sheet.LabSheetA1Sheet.Bath3Blank44_5;

                }

                labSheetDetailNew.TCField1 = null;
                if (!string.IsNullOrWhiteSpace(labSheetAndA1Sheet.LabSheetA1Sheet.TCField1))
                {
                    float temp = 0.0f;
                    if (float.TryParse(labSheetAndA1Sheet.LabSheetA1Sheet.TCField1, out temp))
                    {
                        labSheetDetailNew.TCField1 = temp;
                    }
                }
                labSheetDetailNew.TCLab1 = null;
                if (!string.IsNullOrWhiteSpace(labSheetAndA1Sheet.LabSheetA1Sheet.TCLab1))
                {
                    float temp = 0.0f;
                    if (float.TryParse(labSheetAndA1Sheet.LabSheetA1Sheet.TCLab1, out temp))
                    {
                        labSheetDetailNew.TCLab1 = temp;
                    }
                }
                labSheetDetailNew.TCField2 = null;
                if (!string.IsNullOrWhiteSpace(labSheetAndA1Sheet.LabSheetA1Sheet.TCField2))
                {
                    float temp = 0.0f;
                    if (float.TryParse(labSheetAndA1Sheet.LabSheetA1Sheet.TCField2, out temp))
                    {
                        labSheetDetailNew.TCField2 = temp;
                    }
                }
                labSheetDetailNew.TCLab2 = null;
                if (!string.IsNullOrWhiteSpace(labSheetAndA1Sheet.LabSheetA1Sheet.TCLab2))
                {
                    float temp = 0.0f;
                    if (float.TryParse(labSheetAndA1Sheet.LabSheetA1Sheet.TCLab2, out temp))
                    {
                        labSheetDetailNew.TCLab2 = temp;
                    }
                }
                labSheetDetailNew.TCFirst = null;
                if (!string.IsNullOrWhiteSpace(labSheetAndA1Sheet.LabSheetA1Sheet.TCFirst) && !labSheetAndA1Sheet.LabSheetA1Sheet.TCFirst.Contains("-"))
                {
                    float temp = 0.0f;
                    if (float.TryParse(labSheetAndA1Sheet.LabSheetA1Sheet.TCFirst, out temp))
                    {
                        labSheetDetailNew.TCFirst = temp;
                    }
                }
                labSheetDetailNew.TCAverage = null;
                if (!string.IsNullOrWhiteSpace(labSheetAndA1Sheet.LabSheetA1Sheet.TCAverage) && !labSheetAndA1Sheet.LabSheetA1Sheet.TCAverage.Contains("-"))
                {
                    float temp = 0.0f;
                    if (float.TryParse(labSheetAndA1Sheet.LabSheetA1Sheet.TCAverage, out temp))
                    {
                        labSheetDetailNew.TCAverage = temp;
                    }
                }
                labSheetDetailNew.ControlLot = labSheetAndA1Sheet.LabSheetA1Sheet.ControlLot;
                labSheetDetailNew.Positive35 = labSheetAndA1Sheet.LabSheetA1Sheet.Positive35;
                labSheetDetailNew.NonTarget35 = labSheetAndA1Sheet.LabSheetA1Sheet.NonTarget35;
                labSheetDetailNew.Negative35 = labSheetAndA1Sheet.LabSheetA1Sheet.Negative35;
                labSheetDetailNew.Bath1Positive44_5 = labSheetAndA1Sheet.LabSheetA1Sheet.Bath1Positive44_5;
                labSheetDetailNew.Bath1NonTarget44_5 = labSheetAndA1Sheet.LabSheetA1Sheet.Bath1NonTarget44_5;
                labSheetDetailNew.Bath1Negative44_5 = labSheetAndA1Sheet.LabSheetA1Sheet.Bath1Negative44_5;
                labSheetDetailNew.Blank35 = labSheetAndA1Sheet.LabSheetA1Sheet.Blank35;
                labSheetDetailNew.Bath1Blank44_5 = labSheetAndA1Sheet.LabSheetA1Sheet.Bath1Blank44_5;
                labSheetDetailNew.Lot35 = labSheetAndA1Sheet.LabSheetA1Sheet.Lot35;
                labSheetDetailNew.Lot44_5 = labSheetAndA1Sheet.LabSheetA1Sheet.Lot44_5;
                labSheetDetailNew.RunComment = labSheetAndA1Sheet.LabSheetA1Sheet.RunComment;
                labSheetDetailNew.RunWeatherComment = labSheetAndA1Sheet.LabSheetA1Sheet.RunWeatherComment;
                labSheetDetailNew.SampleBottleLotNumber = labSheetAndA1Sheet.LabSheetA1Sheet.SampleBottleLotNumber;
                labSheetDetailNew.SalinitiesReadBy = labSheetAndA1Sheet.LabSheetA1Sheet.SalinitiesReadBy;

                // SalinitiesReadDate
                DateTime SalinitiesReadDate = new DateTime();
                int SalinitiesReadYear = (string.IsNullOrWhiteSpace(labSheetAndA1Sheet.LabSheetA1Sheet.SalinitiesReadYear) ? 1900 : int.Parse(labSheetAndA1Sheet.LabSheetA1Sheet.SalinitiesReadYear));
                int SalinitiesReadMonth = (string.IsNullOrWhiteSpace(labSheetAndA1Sheet.LabSheetA1Sheet.SalinitiesReadMonth) ? 1 : int.Parse(labSheetAndA1Sheet.LabSheetA1Sheet.SalinitiesReadMonth));
                int SalinitiesReadDay = (string.IsNullOrWhiteSpace(labSheetAndA1Sheet.LabSheetA1Sheet.SalinitiesReadDay) ? 1 : int.Parse(labSheetAndA1Sheet.LabSheetA1Sheet.SalinitiesReadDay));
                if (SalinitiesReadYear != 1900)
                {
                    SalinitiesReadDate = new DateTime(SalinitiesReadYear, SalinitiesReadMonth, SalinitiesReadDay);
                }

                labSheetDetailNew.SalinitiesReadDate = SalinitiesReadDate;
                labSheetDetailNew.ResultsReadBy = labSheetAndA1Sheet.LabSheetA1Sheet.ResultsReadBy;

                // ResultsReadDate
                DateTime ResultsReadDate = new DateTime();
                int ResultsReadYear = (string.IsNullOrWhiteSpace(labSheetAndA1Sheet.LabSheetA1Sheet.ResultsReadYear) ? 1900 : int.Parse(labSheetAndA1Sheet.LabSheetA1Sheet.ResultsReadYear));
                int ResultsReadMonth = (string.IsNullOrWhiteSpace(labSheetAndA1Sheet.LabSheetA1Sheet.ResultsReadMonth) ? 1 : int.Parse(labSheetAndA1Sheet.LabSheetA1Sheet.ResultsReadMonth));
                int ResultsReadDay = (string.IsNullOrWhiteSpace(labSheetAndA1Sheet.LabSheetA1Sheet.ResultsReadDay) ? 1 : int.Parse(labSheetAndA1Sheet.LabSheetA1Sheet.ResultsReadDay));
                if (ResultsReadYear != 1900)
                {
                    ResultsReadDate = new DateTime(ResultsReadYear, ResultsReadMonth, ResultsReadDay);
                }

                labSheetDetailNew.ResultsReadDate = ResultsReadDate;
                labSheetDetailNew.ResultsRecordedBy = labSheetAndA1Sheet.LabSheetA1Sheet.ResultsRecordedBy;

                // ResultsRecordedDate
                DateTime ResultsRecordedDate = new DateTime();
                int ResultsRecordedYear = (string.IsNullOrWhiteSpace(labSheetAndA1Sheet.LabSheetA1Sheet.ResultsRecordedYear) ? 1900 : int.Parse(labSheetAndA1Sheet.LabSheetA1Sheet.ResultsRecordedYear));
                int ResultsRecordedMonth = (string.IsNullOrWhiteSpace(labSheetAndA1Sheet.LabSheetA1Sheet.ResultsRecordedMonth) ? 1 : int.Parse(labSheetAndA1Sheet.LabSheetA1Sheet.ResultsRecordedMonth));
                int ResultsRecordedDay = (string.IsNullOrWhiteSpace(labSheetAndA1Sheet.LabSheetA1Sheet.ResultsRecordedDay) ? 1 : int.Parse(labSheetAndA1Sheet.LabSheetA1Sheet.ResultsRecordedDay));
                if (ResultsRecordedYear != 1900)
                {
                    ResultsRecordedDate = new DateTime(ResultsRecordedYear, ResultsRecordedMonth, ResultsRecordedDay);
                }

                labSheetDetailNew.ResultsRecordedDate = ResultsRecordedDate;

                labSheetDetailNew.DailyDuplicateRLog = 0.0f;
                if (!string.IsNullOrWhiteSpace(labSheetAndA1Sheet.LabSheetA1Sheet.DailyDuplicateRLog) && !labSheetAndA1Sheet.LabSheetA1Sheet.DailyDuplicateRLog.StartsWith("N"))
                {
                    try
                    {
                        labSheetDetailNew.DailyDuplicateRLog = float.Parse(labSheetAndA1Sheet.LabSheetA1Sheet.DailyDuplicateRLog);
                    }
                    catch (Exception)
                    {
                        // nothing
                    }
                }

                labSheetDetailNew.DailyDuplicatePrecisionCriteria = null;
                labSheetDetailNew.DailyDuplicateRLog = null;
                labSheetDetailNew.DailyDuplicateAcceptable = null;
                labSheetDetailNew.IntertechDuplicatePrecisionCriteria = null;
                labSheetDetailNew.IntertechDuplicateRLog = null;
                labSheetDetailNew.IntertechDuplicateAcceptable = null;
                labSheetDetailNew.IntertechReadAcceptable = null;

                if (!string.IsNullOrWhiteSpace(labSheetAndA1Sheet.LabSheetA1Sheet.DailyDuplicatePrecisionCriteria))
                {
                    labSheetDetailNew.DailyDuplicatePrecisionCriteria = float.Parse(labSheetAndA1Sheet.LabSheetA1Sheet.DailyDuplicatePrecisionCriteria);
                }

                if (!string.IsNullOrWhiteSpace(labSheetAndA1Sheet.LabSheetA1Sheet.DailyDuplicateRLog))
                {
                    if (!labSheetAndA1Sheet.LabSheetA1Sheet.DailyDuplicateRLog.StartsWith("N"))
                    {
                        labSheetDetailNew.DailyDuplicateRLog = float.Parse(labSheetAndA1Sheet.LabSheetA1Sheet.DailyDuplicateRLog);
                    }
                }

                if (!string.IsNullOrWhiteSpace(labSheetAndA1Sheet.LabSheetA1Sheet.DailyDuplicateAcceptableOrUnacceptable))
                {
                    labSheetDetailNew.DailyDuplicateAcceptable = (labSheetAndA1Sheet.LabSheetA1Sheet.DailyDuplicateAcceptableOrUnacceptable != "Acceptable" ? false : true);
                }

                if (!string.IsNullOrWhiteSpace(labSheetAndA1Sheet.LabSheetA1Sheet.IntertechDuplicatePrecisionCriteria))
                {
                    labSheetDetailNew.IntertechDuplicatePrecisionCriteria = float.Parse(labSheetAndA1Sheet.LabSheetA1Sheet.IntertechDuplicatePrecisionCriteria);
                }

                if (!string.IsNullOrWhiteSpace(labSheetAndA1Sheet.LabSheetA1Sheet.IntertechDuplicateRLog))
                {
                    if (!labSheetAndA1Sheet.LabSheetA1Sheet.IntertechDuplicateRLog.StartsWith("N"))
                    {
                        labSheetDetailNew.IntertechDuplicateRLog = float.Parse(labSheetAndA1Sheet.LabSheetA1Sheet.IntertechDuplicateRLog);
                    }
                }

                if (!string.IsNullOrWhiteSpace(labSheetAndA1Sheet.LabSheetA1Sheet.IntertechDuplicateAcceptableOrUnacceptable))
                {
                    labSheetDetailNew.IntertechDuplicateAcceptable = (labSheetAndA1Sheet.LabSheetA1Sheet.IntertechDuplicateAcceptableOrUnacceptable != "Acceptable" ? false : true);
                }

                if (!string.IsNullOrWhiteSpace(labSheetAndA1Sheet.LabSheetA1Sheet.IntertechReadAcceptableOrUnacceptable))
                {
                    labSheetDetailNew.IntertechReadAcceptable = (labSheetAndA1Sheet.LabSheetA1Sheet.IntertechReadAcceptableOrUnacceptable != "Acceptable" ? false : true);
                }
            }

            LabSheetDetailService labSheetDetailService = new LabSheetDetailService(LanguageEnum.en, _User);

            LabSheetDetail labSheetDetailExist = labSheetDetailService.GetLabSheetDetailExistDB(labSheetDetailNew);
            if (!string.IsNullOrWhiteSpace(labSheetDetailExist.Error))
            {
                labSheetDetailExist = labSheetDetailService.PostAddLabSheetDetailDB(labSheetDetailNew);
                if (!string.IsNullOrWhiteSpace(labSheetDetailExist.Error))
                {
                    richTextBoxStatus.AppendText("Lab sheet detail could not be loaded to the local DB. Error [" + labSheetDetailExist.Error + "]\r\n");
                    return labSheetDetailExist.Error;
                }
            }
            else
            {
                labSheetDetailNew.LabSheetDetailID = labSheetDetailExist.LabSheetDetailID;
                labSheetDetailExist = labSheetDetailService.PostUpdateLabSheetDetailDB(labSheetDetailNew);
                if (!string.IsNullOrWhiteSpace(labSheetDetailExist.Error))
                {
                    richTextBoxStatus.AppendText("Lab sheet detail could not be loaded to the local DB. Error [" + labSheetDetailExist.Error + "]\r\n");
                    return labSheetDetailExist.Error;
                }
            }

            retStr = UploadLabSheetTubeMPNDetailInDB(labSheetDetailExist.LabSheetDetailID, labSheetAndA1Sheet.LabSheetA1Sheet.LabSheetA1MeasurementList);
            if (!string.IsNullOrWhiteSpace(retStr))
            {
                richTextBoxStatus.AppendText("Lab sheet tube and MPN detail could not be loaded to the local DB. Error [" + retStr + "]\r\n");
                return retStr;
            }

            return retStr;
        }
        public string UploadLabSheetTubeMPNDetailInDB(int LabSheetDetailID, List<LabSheetA1Measurement> labSheetA1MeasurementList)
        {
            string retStr = "";

            LabSheetTubeMPNDetailService labSheetTubeMPNDetailService = new LabSheetTubeMPNDetailService(LanguageEnum.en, _User);

            int Ordinal = 0;
            foreach (LabSheetA1Measurement labSheetA1Measurement in labSheetA1MeasurementList)
            {
                LabSheetTubeMPNDetail labSheetTubeMPNDetailNew = new LabSheetTubeMPNDetail();
                labSheetTubeMPNDetailNew.LabSheetDetailID = LabSheetDetailID;
                labSheetTubeMPNDetailNew.Ordinal = Ordinal;
                labSheetTubeMPNDetailNew.MWQMSiteTVItemID = labSheetA1Measurement.TVItemID;
                labSheetTubeMPNDetailNew.SampleDateTime = labSheetA1Measurement.Time;
                labSheetTubeMPNDetailNew.MPN = labSheetA1Measurement.MPN;
                labSheetTubeMPNDetailNew.Tube10 = labSheetA1Measurement.Tube10;
                labSheetTubeMPNDetailNew.Tube1_0 = labSheetA1Measurement.Tube1_0;
                labSheetTubeMPNDetailNew.Tube0_1 = labSheetA1Measurement.Tube0_1;
                labSheetTubeMPNDetailNew.Salinity = labSheetA1Measurement.Salinity;
                labSheetTubeMPNDetailNew.Temperature = labSheetA1Measurement.Temperature;
                labSheetTubeMPNDetailNew.ProcessedBy = labSheetA1Measurement.ProcessedBy;
                labSheetTubeMPNDetailNew.SampleType = (SampleTypeEnum)labSheetA1Measurement.SampleType;
                labSheetTubeMPNDetailNew.SiteComment = labSheetA1Measurement.SiteComment;

                LabSheetTubeMPNDetail labSheetTubeMPNDetailExist = labSheetTubeMPNDetailService.GetLabSheetTubeMPNDetailExistDB(labSheetTubeMPNDetailNew);
                if (!string.IsNullOrWhiteSpace(labSheetTubeMPNDetailExist.Error))
                {
                    labSheetTubeMPNDetailExist = labSheetTubeMPNDetailService.PostAddLabSheetTubeMPNDetailDB(labSheetTubeMPNDetailNew);
                    if (!string.IsNullOrWhiteSpace(labSheetTubeMPNDetailExist.Error))
                        return labSheetTubeMPNDetailExist.Error;
                }
                else
                {
                    labSheetTubeMPNDetailNew.LabSheetTubeMPNDetailID = labSheetTubeMPNDetailExist.LabSheetTubeMPNDetailID;
                    labSheetTubeMPNDetailExist = labSheetTubeMPNDetailService.PostUpdateLabSheetTubeMPNDetailDB(labSheetTubeMPNDetailNew);
                    if (!string.IsNullOrWhiteSpace(labSheetTubeMPNDetailExist.Error))
                        return labSheetTubeMPNDetailExist.Error;
                }

                Ordinal += 1;
            }

            return retStr;
        }
        public void GetNextLabSheet()
        {
            string retStr = "";
            try
            {
                using (WebClient webClient = new WebClient())
                {
                    WebProxy webProxy = new WebProxy();
                    webClient.Proxy = webProxy;

                    string FullLabSheetText = webClient.DownloadString(new Uri("http://cssplabsheet.azurewebsites.net/GetNextLabSheet.aspx"));

                    if (FullLabSheetText.Length > 0)
                    {
                        int posStart = FullLabSheetText.IndexOf("OtherServerLabSheetID|||||[") + 27;
                        int posEnd = FullLabSheetText.IndexOf("]", posStart);
                        string OtherServerLabSheetIDTxt = FullLabSheetText.Substring(posStart, posEnd - posStart);
                        int OtherServerLabSheetID = int.Parse(OtherServerLabSheetIDTxt);

                        posStart = FullLabSheetText.IndexOf("SubsectorTVItemID|||||[") + 23;
                        posEnd = FullLabSheetText.IndexOf("]", posStart);
                        string SubsectorTVItemIDTxt = FullLabSheetText.Substring(posStart, posEnd - posStart);
                        int SubsectorTVItemID = int.Parse(SubsectorTVItemIDTxt);

                        TVItem tvItemSubsector = new TVItem();
                        TVItem tvItemCountry = new TVItem();
                        TVItem tvItemProvince = new TVItem();
                        LabSheetAndA1Sheet labSheetAndA1Sheet = new LabSheetAndA1Sheet();
                        using (TransactionScope ts = new TransactionScope())
                        {
                            LabSheetService labSheetService = new LabSheetService(LanguageEnum.en, _User);

                            LabSheet labSheetRet = labSheetService.AddOrUpdateLabSheetDB(FullLabSheetText);
                            if (!string.IsNullOrWhiteSpace(labSheetRet.Error))
                            {
                                richTextBoxStatus.AppendText("Lab sheet adding error OtherServerLabSheetID [" + OtherServerLabSheetID.ToString() + "]" + labSheetRet.Error + "]\r\n");
                                return;
                            }

                            TVItemService tvItemService = new TVItemService(LanguageEnum.en, _User);

                            tvItemSubsector = tvItemService.GetTVItemModelWithTVItemIDDB(SubsectorTVItemID);
                            if (!string.IsNullOrWhiteSpace(tvItemSubsector.Error))
                            {
                                richTextBoxStatus.AppendText("Lab sheet parsing error OtherServerLabSheetID [" + OtherServerLabSheetID.ToString() + "]" + tvItemSubsector.Error + "]\r\n");
                                return;
                            }

                            List<TVItem> tvItemList = tvItemService.GetParentsTVItemList(tvItemSubsector.TVPath);
                            foreach (TVItem tvItem in tvItemList)
                            {
                                if (tvItem.TVType == TVTypeEnum.Province)
                                {
                                    tvItemProvince = tvItem;
                                }
                                if (tvItem.TVType == TVTypeEnum.Country)
                                {
                                    tvItemCountry = tvItem;
                                }
                            }

                            labSheetAndA1Sheet.LabSheet = labSheetRet;
                            labSheetAndA1Sheet.LabSheetA1Sheet = labSheetService.ParseLabSheetA1WithLabSheetID(labSheetRet.LabSheetID);
                            if (!string.IsNullOrWhiteSpace(labSheetAndA1Sheet.LabSheetA1Sheet.Error))
                            {
                                richTextBoxStatus.AppendText("Lab sheet parsing error OtherServerLabSheetID [" + OtherServerLabSheetID.ToString() + "]" + labSheetAndA1Sheet.LabSheetA1Sheet.Error + "]\r\n");
                                richTextBoxStatus.AppendText("Full Lab Sheet Text below\r\n");
                                richTextBoxStatus.AppendText("---------------- Start of full lab sheet text -----------\r\n");
                                richTextBoxStatus.AppendText(FullLabSheetText);
                                richTextBoxStatus.AppendText("---------------- End of full lab sheet text -----------\r\n");
                                retStr = UpdateOtherServerWithOtherServerLabSheetIDAndLabSheetStatus(OtherServerLabSheetID, LabSheetStatusEnum.Error);
                                if (!string.IsNullOrWhiteSpace(retStr))
                                {
                                    richTextBoxStatus.AppendText("Error updating other server lab sheet [" + retStr + "]");
                                }
                                return;
                            }

                            string retStr2 = UploadLabSheetDetailInDB(labSheetAndA1Sheet);
                            if (!string.IsNullOrWhiteSpace(retStr2))
                            {
                                // Error message already sent to richTextboxStatus
                                retStr = UpdateOtherServerWithOtherServerLabSheetIDAndLabSheetStatus(OtherServerLabSheetID, LabSheetStatusEnum.Error);
                                if (!string.IsNullOrWhiteSpace(retStr))
                                {
                                    richTextBoxStatus.AppendText("Error updating other server lab sheet [" + retStr + "]");
                                }
                                return;
                            }

                            ts.Complete();
                        }

                        string href = "http://wmon01dtchlebl2/csspwebtools/en-CA/#!View/" + (tvItemCountry.TVText + "-" + tvItemProvince.TVText).Replace(" ", "-") + "|||" + tvItemProvince.TVItemID.ToString() + "|||010003030200000000000000000000";

                        if (labSheetAndA1Sheet.LabSheetA1Sheet.LabSheetA1MeasurementList.Where(c => c.MPN != null && c.MPN >= MPNLimitForEmail).Any())
                        {
                            SendNewLabSheetEmailBigMPN(href, tvItemProvince, tvItemSubsector, labSheetAndA1Sheet);
                        }
                        else
                        {
                            SendNewLabSheetEmail(href, tvItemProvince, tvItemSubsector, labSheetAndA1Sheet);
                        }

                        retStr = UpdateOtherServerWithOtherServerLabSheetIDAndLabSheetStatus(OtherServerLabSheetID, LabSheetStatusEnum.Transferred);
                        if (!string.IsNullOrWhiteSpace(retStr))
                        {
                            richTextBoxStatus.AppendText("Error updating other server lab sheet [" + retStr + "]");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                string errrrrr = ex.Message;
            }
        }
        #endregion Functions public

        #region Functions private
        #endregion Functions private
    }
}
