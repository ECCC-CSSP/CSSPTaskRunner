using CSSPEnums;
using CSSPModels;
using CSSPTaskRunnerHelper.Resources;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Principal;
using System.Text;
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
        public string UploadLabSheetDetailInDB(LabSheetModelAndA1Sheet labSheetModelAndA1Sheet)
        {
            string retStr = "";

            // Filling LabSheetDetailModel
            LabSheetDetailModel labSheetDetailModelNew = new LabSheetDetailModel();

            labSheetDetailModelNew.LabSheetID = labSheetModelAndA1Sheet.LabSheetModel.LabSheetID;
            labSheetDetailModelNew.SamplingPlanID = labSheetModelAndA1Sheet.LabSheetModel.SamplingPlanID;
            labSheetDetailModelNew.SubsectorTVItemID = labSheetModelAndA1Sheet.LabSheetModel.SubsectorTVItemID;
            labSheetDetailModelNew.Version = labSheetModelAndA1Sheet.LabSheetA1Sheet.Version;
            labSheetDetailModelNew.IncludeLaboratoryQAQC = labSheetModelAndA1Sheet.LabSheetA1Sheet.IncludeLaboratoryQAQC;

            // RunDate
            DateTime RunDate = new DateTime();
            int RunYear = (string.IsNullOrWhiteSpace(labSheetModelAndA1Sheet.LabSheetA1Sheet.RunYear) ? 1900 : int.Parse(labSheetModelAndA1Sheet.LabSheetA1Sheet.RunYear));
            int RunMonth = (string.IsNullOrWhiteSpace(labSheetModelAndA1Sheet.LabSheetA1Sheet.RunMonth) ? 1 : int.Parse(labSheetModelAndA1Sheet.LabSheetA1Sheet.RunMonth));
            int RunDay = (string.IsNullOrWhiteSpace(labSheetModelAndA1Sheet.LabSheetA1Sheet.RunDay) ? 1 : int.Parse(labSheetModelAndA1Sheet.LabSheetA1Sheet.RunDay));
            if (RunYear != 1900)
            {
                RunDate = new DateTime(RunYear, RunMonth, RunDay);
            }

            labSheetDetailModelNew.RunDate = RunDate;

            labSheetDetailModelNew.Tides = labSheetModelAndA1Sheet.LabSheetA1Sheet.Tides;
            labSheetDetailModelNew.SampleCrewInitials = labSheetModelAndA1Sheet.LabSheetA1Sheet.SampleCrewInitials;

            if (labSheetModelAndA1Sheet.LabSheetA1Sheet.IncludeLaboratoryQAQC)
            {
                // IncubationBath1StartDate
                DateTime IncubationBath1StartDate = new DateTime();
                int IncubationBath1StartYear = (string.IsNullOrWhiteSpace(labSheetModelAndA1Sheet.LabSheetA1Sheet.RunYear) ? 1900 : int.Parse(labSheetModelAndA1Sheet.LabSheetA1Sheet.RunYear));
                int IncubationBath1StartMonth = (string.IsNullOrWhiteSpace(labSheetModelAndA1Sheet.LabSheetA1Sheet.RunMonth) ? 1 : int.Parse(labSheetModelAndA1Sheet.LabSheetA1Sheet.RunMonth));
                int IncubationBath1StartDay = (string.IsNullOrWhiteSpace(labSheetModelAndA1Sheet.LabSheetA1Sheet.RunDay) ? 1 : int.Parse(labSheetModelAndA1Sheet.LabSheetA1Sheet.RunDay));
                int IncubationBath1StartHour = 0;
                int IncubationBath1StartMinute = 0;
                if (labSheetModelAndA1Sheet.LabSheetA1Sheet.IncubationBath1StartTime.Length == 5)
                {
                    IncubationBath1StartHour = int.Parse(labSheetModelAndA1Sheet.LabSheetA1Sheet.IncubationBath1StartTime.Substring(0, 2));
                    IncubationBath1StartMinute = int.Parse(labSheetModelAndA1Sheet.LabSheetA1Sheet.IncubationBath1StartTime.Substring(3, 2));
                }
                if (IncubationBath1StartYear != 1900)
                {
                    IncubationBath1StartDate = new DateTime(IncubationBath1StartYear, IncubationBath1StartMonth, IncubationBath1StartDay, IncubationBath1StartHour, IncubationBath1StartMinute, 0);
                }

                labSheetDetailModelNew.IncubationBath1StartTime = IncubationBath1StartDate;

                // IncubationBath1EndDate
                DateTime IncubationBath1EndDate = new DateTime();
                int IncubationBath1EndYear = (string.IsNullOrWhiteSpace(labSheetModelAndA1Sheet.LabSheetA1Sheet.RunYear) ? 1900 : int.Parse(labSheetModelAndA1Sheet.LabSheetA1Sheet.RunYear));
                int IncubationBath1EndMonth = (string.IsNullOrWhiteSpace(labSheetModelAndA1Sheet.LabSheetA1Sheet.RunMonth) ? 1 : int.Parse(labSheetModelAndA1Sheet.LabSheetA1Sheet.RunMonth));
                int IncubationBath1EndDay = (string.IsNullOrWhiteSpace(labSheetModelAndA1Sheet.LabSheetA1Sheet.RunDay) ? 1 : int.Parse(labSheetModelAndA1Sheet.LabSheetA1Sheet.RunDay));
                int IncubationBath1EndHour = 0;
                int IncubationBath1EndMinute = 0;
                if (labSheetModelAndA1Sheet.LabSheetA1Sheet.IncubationBath1EndTime.Length == 5)
                {
                    IncubationBath1EndHour = int.Parse(labSheetModelAndA1Sheet.LabSheetA1Sheet.IncubationBath1EndTime.Substring(0, 2));
                    IncubationBath1EndMinute = int.Parse(labSheetModelAndA1Sheet.LabSheetA1Sheet.IncubationBath1EndTime.Substring(3, 2));
                }
                if (IncubationBath1EndYear != 1900)
                {
                    IncubationBath1EndDate = new DateTime(IncubationBath1EndYear, IncubationBath1EndMonth, IncubationBath1EndDay, IncubationBath1EndHour, IncubationBath1EndMinute, 0).AddDays(1);
                }

                labSheetDetailModelNew.IncubationBath1EndTime = IncubationBath1EndDate;

                // IncubationBath1TimeCalculated_minutes 
                int IncubationBath1TimeCalculated_minutes = 0;
                if (labSheetModelAndA1Sheet.LabSheetA1Sheet.IncubationBath1TimeCalculated.Length == 5)
                {
                    IncubationBath1TimeCalculated_minutes = int.Parse(labSheetModelAndA1Sheet.LabSheetA1Sheet.IncubationBath1TimeCalculated.Substring(0, 2)) * 60;
                    IncubationBath1TimeCalculated_minutes += int.Parse(labSheetModelAndA1Sheet.LabSheetA1Sheet.IncubationBath1TimeCalculated.Substring(3, 2));
                }

                labSheetDetailModelNew.IncubationBath1TimeCalculated_minutes = IncubationBath1TimeCalculated_minutes;

                labSheetDetailModelNew.WaterBath1 = labSheetModelAndA1Sheet.LabSheetA1Sheet.WaterBath1;

                labSheetDetailModelNew.WaterBathCount = labSheetModelAndA1Sheet.LabSheetA1Sheet.WaterBathCount;

                if (labSheetDetailModelNew.WaterBathCount > 1)
                {
                    // IncubationBath2StartDate
                    DateTime IncubationBath2StartDate = new DateTime();
                    int IncubationBath2StartYear = (string.IsNullOrWhiteSpace(labSheetModelAndA1Sheet.LabSheetA1Sheet.RunYear) ? 1900 : int.Parse(labSheetModelAndA1Sheet.LabSheetA1Sheet.RunYear));
                    int IncubationBath2StartMonth = (string.IsNullOrWhiteSpace(labSheetModelAndA1Sheet.LabSheetA1Sheet.RunMonth) ? 1 : int.Parse(labSheetModelAndA1Sheet.LabSheetA1Sheet.RunMonth));
                    int IncubationBath2StartDay = (string.IsNullOrWhiteSpace(labSheetModelAndA1Sheet.LabSheetA1Sheet.RunDay) ? 1 : int.Parse(labSheetModelAndA1Sheet.LabSheetA1Sheet.RunDay));
                    int IncubationBath2StartHour = 0;
                    int IncubationBath2StartMinute = 0;
                    if (labSheetModelAndA1Sheet.LabSheetA1Sheet.IncubationBath2StartTime.Length == 5)
                    {
                        IncubationBath2StartHour = int.Parse(labSheetModelAndA1Sheet.LabSheetA1Sheet.IncubationBath2StartTime.Substring(0, 2));
                        IncubationBath2StartMinute = int.Parse(labSheetModelAndA1Sheet.LabSheetA1Sheet.IncubationBath2StartTime.Substring(3, 2));
                    }
                    if (IncubationBath2StartYear != 1900)
                    {
                        IncubationBath2StartDate = new DateTime(IncubationBath2StartYear, IncubationBath2StartMonth, IncubationBath2StartDay, IncubationBath2StartHour, IncubationBath2StartMinute, 0);
                    }

                    labSheetDetailModelNew.IncubationBath2StartTime = IncubationBath2StartDate;

                    // IncubationBath2EndDate
                    DateTime IncubationBath2EndDate = new DateTime();
                    int IncubationBath2EndYear = (string.IsNullOrWhiteSpace(labSheetModelAndA1Sheet.LabSheetA1Sheet.RunYear) ? 1900 : int.Parse(labSheetModelAndA1Sheet.LabSheetA1Sheet.RunYear));
                    int IncubationBath2EndMonth = (string.IsNullOrWhiteSpace(labSheetModelAndA1Sheet.LabSheetA1Sheet.RunMonth) ? 1 : int.Parse(labSheetModelAndA1Sheet.LabSheetA1Sheet.RunMonth));
                    int IncubationBath2EndDay = (string.IsNullOrWhiteSpace(labSheetModelAndA1Sheet.LabSheetA1Sheet.RunDay) ? 1 : int.Parse(labSheetModelAndA1Sheet.LabSheetA1Sheet.RunDay));
                    int IncubationBath2EndHour = 0;
                    int IncubationBath2EndMinute = 0;
                    if (labSheetModelAndA1Sheet.LabSheetA1Sheet.IncubationBath2EndTime.Length == 5)
                    {
                        IncubationBath2EndHour = int.Parse(labSheetModelAndA1Sheet.LabSheetA1Sheet.IncubationBath2EndTime.Substring(0, 2));
                        IncubationBath2EndMinute = int.Parse(labSheetModelAndA1Sheet.LabSheetA1Sheet.IncubationBath2EndTime.Substring(3, 2));
                    }
                    if (IncubationBath2EndYear != 1900)
                    {
                        IncubationBath2EndDate = new DateTime(IncubationBath2EndYear, IncubationBath2EndMonth, IncubationBath2EndDay, IncubationBath2EndHour, IncubationBath2EndMinute, 0).AddDays(1);
                    }

                    labSheetDetailModelNew.IncubationBath2EndTime = IncubationBath2EndDate;

                    // IncubationBath2TimeCalculated_minutes 
                    int IncubationBath2TimeCalculated_minutes = 0;
                    if (labSheetModelAndA1Sheet.LabSheetA1Sheet.IncubationBath2TimeCalculated.Length == 5)
                    {
                        IncubationBath2TimeCalculated_minutes = int.Parse(labSheetModelAndA1Sheet.LabSheetA1Sheet.IncubationBath2TimeCalculated.Substring(0, 2)) * 60;
                        IncubationBath2TimeCalculated_minutes += int.Parse(labSheetModelAndA1Sheet.LabSheetA1Sheet.IncubationBath2TimeCalculated.Substring(3, 2));
                    }

                    labSheetDetailModelNew.IncubationBath2TimeCalculated_minutes = IncubationBath2TimeCalculated_minutes;

                    labSheetDetailModelNew.WaterBath2 = labSheetModelAndA1Sheet.LabSheetA1Sheet.WaterBath2;
                    labSheetDetailModelNew.Bath2Positive44_5 = labSheetModelAndA1Sheet.LabSheetA1Sheet.Bath2Positive44_5;
                    labSheetDetailModelNew.Bath2NonTarget44_5 = labSheetModelAndA1Sheet.LabSheetA1Sheet.Bath2NonTarget44_5;
                    labSheetDetailModelNew.Bath2Negative44_5 = labSheetModelAndA1Sheet.LabSheetA1Sheet.Bath2Negative44_5;
                    labSheetDetailModelNew.Bath2Blank44_5 = labSheetModelAndA1Sheet.LabSheetA1Sheet.Bath2Blank44_5;

                }

                if (labSheetDetailModelNew.WaterBathCount > 2)
                {
                    // IncubationBath3StartDate
                    DateTime IncubationBath3StartDate = new DateTime();
                    int IncubationBath3StartYear = (string.IsNullOrWhiteSpace(labSheetModelAndA1Sheet.LabSheetA1Sheet.RunYear) ? 1900 : int.Parse(labSheetModelAndA1Sheet.LabSheetA1Sheet.RunYear));
                    int IncubationBath3StartMonth = (string.IsNullOrWhiteSpace(labSheetModelAndA1Sheet.LabSheetA1Sheet.RunMonth) ? 1 : int.Parse(labSheetModelAndA1Sheet.LabSheetA1Sheet.RunMonth));
                    int IncubationBath3StartDay = (string.IsNullOrWhiteSpace(labSheetModelAndA1Sheet.LabSheetA1Sheet.RunDay) ? 1 : int.Parse(labSheetModelAndA1Sheet.LabSheetA1Sheet.RunDay));
                    int IncubationBath3StartHour = 0;
                    int IncubationBath3StartMinute = 0;
                    if (labSheetModelAndA1Sheet.LabSheetA1Sheet.IncubationBath3StartTime.Length == 5)
                    {
                        IncubationBath3StartHour = int.Parse(labSheetModelAndA1Sheet.LabSheetA1Sheet.IncubationBath3StartTime.Substring(0, 2));
                        IncubationBath3StartMinute = int.Parse(labSheetModelAndA1Sheet.LabSheetA1Sheet.IncubationBath3StartTime.Substring(3, 2));
                    }
                    if (IncubationBath3StartYear != 1900)
                    {
                        IncubationBath3StartDate = new DateTime(IncubationBath3StartYear, IncubationBath3StartMonth, IncubationBath3StartDay, IncubationBath3StartHour, IncubationBath3StartMinute, 0);
                    }

                    labSheetDetailModelNew.IncubationBath3StartTime = IncubationBath3StartDate;

                    // IncubationBath3EndDate
                    DateTime IncubationBath3EndDate = new DateTime();
                    int IncubationBath3EndYear = (string.IsNullOrWhiteSpace(labSheetModelAndA1Sheet.LabSheetA1Sheet.RunYear) ? 1900 : int.Parse(labSheetModelAndA1Sheet.LabSheetA1Sheet.RunYear));
                    int IncubationBath3EndMonth = (string.IsNullOrWhiteSpace(labSheetModelAndA1Sheet.LabSheetA1Sheet.RunMonth) ? 1 : int.Parse(labSheetModelAndA1Sheet.LabSheetA1Sheet.RunMonth));
                    int IncubationBath3EndDay = (string.IsNullOrWhiteSpace(labSheetModelAndA1Sheet.LabSheetA1Sheet.RunDay) ? 1 : int.Parse(labSheetModelAndA1Sheet.LabSheetA1Sheet.RunDay));
                    int IncubationBath3EndHour = 0;
                    int IncubationBath3EndMinute = 0;
                    if (labSheetModelAndA1Sheet.LabSheetA1Sheet.IncubationBath3EndTime.Length == 5)
                    {
                        IncubationBath3EndHour = int.Parse(labSheetModelAndA1Sheet.LabSheetA1Sheet.IncubationBath3EndTime.Substring(0, 2));
                        IncubationBath3EndMinute = int.Parse(labSheetModelAndA1Sheet.LabSheetA1Sheet.IncubationBath3EndTime.Substring(3, 2));
                    }
                    if (IncubationBath3EndYear != 1900)
                    {
                        IncubationBath3EndDate = new DateTime(IncubationBath3EndYear, IncubationBath3EndMonth, IncubationBath3EndDay, IncubationBath3EndHour, IncubationBath3EndMinute, 0).AddDays(1);
                    }

                    labSheetDetailModelNew.IncubationBath3EndTime = IncubationBath3EndDate;

                    // IncubationBath3TimeCalculated_minutes 
                    int IncubationBath3TimeCalculated_minutes = 0;
                    if (labSheetModelAndA1Sheet.LabSheetA1Sheet.IncubationBath3TimeCalculated.Length == 5)
                    {
                        IncubationBath3TimeCalculated_minutes = int.Parse(labSheetModelAndA1Sheet.LabSheetA1Sheet.IncubationBath3TimeCalculated.Substring(0, 2)) * 60;
                        IncubationBath3TimeCalculated_minutes += int.Parse(labSheetModelAndA1Sheet.LabSheetA1Sheet.IncubationBath3TimeCalculated.Substring(3, 2));
                    }

                    labSheetDetailModelNew.IncubationBath3TimeCalculated_minutes = IncubationBath3TimeCalculated_minutes;

                    labSheetDetailModelNew.WaterBath3 = labSheetModelAndA1Sheet.LabSheetA1Sheet.WaterBath3;
                    labSheetDetailModelNew.Bath3Positive44_5 = labSheetModelAndA1Sheet.LabSheetA1Sheet.Bath3Positive44_5;
                    labSheetDetailModelNew.Bath3NonTarget44_5 = labSheetModelAndA1Sheet.LabSheetA1Sheet.Bath3NonTarget44_5;
                    labSheetDetailModelNew.Bath3Negative44_5 = labSheetModelAndA1Sheet.LabSheetA1Sheet.Bath3Negative44_5;
                    labSheetDetailModelNew.Bath3Blank44_5 = labSheetModelAndA1Sheet.LabSheetA1Sheet.Bath3Blank44_5;

                }

                labSheetDetailModelNew.TCField1 = null;
                if (!string.IsNullOrWhiteSpace(labSheetModelAndA1Sheet.LabSheetA1Sheet.TCField1))
                {
                    float temp = 0.0f;
                    if (float.TryParse(labSheetModelAndA1Sheet.LabSheetA1Sheet.TCField1, out temp))
                    {
                        labSheetDetailModelNew.TCField1 = temp;
                    }
                }
                labSheetDetailModelNew.TCLab1 = null;
                if (!string.IsNullOrWhiteSpace(labSheetModelAndA1Sheet.LabSheetA1Sheet.TCLab1))
                {
                    float temp = 0.0f;
                    if (float.TryParse(labSheetModelAndA1Sheet.LabSheetA1Sheet.TCLab1, out temp))
                    {
                        labSheetDetailModelNew.TCLab1 = temp;
                    }
                }
                labSheetDetailModelNew.TCField2 = null;
                if (!string.IsNullOrWhiteSpace(labSheetModelAndA1Sheet.LabSheetA1Sheet.TCField2))
                {
                    float temp = 0.0f;
                    if (float.TryParse(labSheetModelAndA1Sheet.LabSheetA1Sheet.TCField2, out temp))
                    {
                        labSheetDetailModelNew.TCField2 = temp;
                    }
                }
                labSheetDetailModelNew.TCLab2 = null;
                if (!string.IsNullOrWhiteSpace(labSheetModelAndA1Sheet.LabSheetA1Sheet.TCLab2))
                {
                    float temp = 0.0f;
                    if (float.TryParse(labSheetModelAndA1Sheet.LabSheetA1Sheet.TCLab2, out temp))
                    {
                        labSheetDetailModelNew.TCLab2 = temp;
                    }
                }
                labSheetDetailModelNew.TCFirst = null;
                if (!string.IsNullOrWhiteSpace(labSheetModelAndA1Sheet.LabSheetA1Sheet.TCFirst) && !labSheetModelAndA1Sheet.LabSheetA1Sheet.TCFirst.Contains("-"))
                {
                    float temp = 0.0f;
                    if (float.TryParse(labSheetModelAndA1Sheet.LabSheetA1Sheet.TCFirst, out temp))
                    {
                        labSheetDetailModelNew.TCFirst = temp;
                    }
                }
                labSheetDetailModelNew.TCAverage = null;
                if (!string.IsNullOrWhiteSpace(labSheetModelAndA1Sheet.LabSheetA1Sheet.TCAverage) && !labSheetModelAndA1Sheet.LabSheetA1Sheet.TCAverage.Contains("-"))
                {
                    float temp = 0.0f;
                    if (float.TryParse(labSheetModelAndA1Sheet.LabSheetA1Sheet.TCAverage, out temp))
                    {
                        labSheetDetailModelNew.TCAverage = temp;
                    }
                }
                labSheetDetailModelNew.ControlLot = labSheetModelAndA1Sheet.LabSheetA1Sheet.ControlLot;
                labSheetDetailModelNew.Positive35 = labSheetModelAndA1Sheet.LabSheetA1Sheet.Positive35;
                labSheetDetailModelNew.NonTarget35 = labSheetModelAndA1Sheet.LabSheetA1Sheet.NonTarget35;
                labSheetDetailModelNew.Negative35 = labSheetModelAndA1Sheet.LabSheetA1Sheet.Negative35;
                labSheetDetailModelNew.Bath1Positive44_5 = labSheetModelAndA1Sheet.LabSheetA1Sheet.Bath1Positive44_5;
                labSheetDetailModelNew.Bath1NonTarget44_5 = labSheetModelAndA1Sheet.LabSheetA1Sheet.Bath1NonTarget44_5;
                labSheetDetailModelNew.Bath1Negative44_5 = labSheetModelAndA1Sheet.LabSheetA1Sheet.Bath1Negative44_5;
                labSheetDetailModelNew.Blank35 = labSheetModelAndA1Sheet.LabSheetA1Sheet.Blank35;
                labSheetDetailModelNew.Bath1Blank44_5 = labSheetModelAndA1Sheet.LabSheetA1Sheet.Bath1Blank44_5;
                labSheetDetailModelNew.Lot35 = labSheetModelAndA1Sheet.LabSheetA1Sheet.Lot35;
                labSheetDetailModelNew.Lot44_5 = labSheetModelAndA1Sheet.LabSheetA1Sheet.Lot44_5;
                labSheetDetailModelNew.RunComment = labSheetModelAndA1Sheet.LabSheetA1Sheet.RunComment;
                labSheetDetailModelNew.RunWeatherComment = labSheetModelAndA1Sheet.LabSheetA1Sheet.RunWeatherComment;
                labSheetDetailModelNew.SampleBottleLotNumber = labSheetModelAndA1Sheet.LabSheetA1Sheet.SampleBottleLotNumber;
                labSheetDetailModelNew.SalinitiesReadBy = labSheetModelAndA1Sheet.LabSheetA1Sheet.SalinitiesReadBy;

                // SalinitiesReadDate
                DateTime SalinitiesReadDate = new DateTime();
                int SalinitiesReadYear = (string.IsNullOrWhiteSpace(labSheetModelAndA1Sheet.LabSheetA1Sheet.SalinitiesReadYear) ? 1900 : int.Parse(labSheetModelAndA1Sheet.LabSheetA1Sheet.SalinitiesReadYear));
                int SalinitiesReadMonth = (string.IsNullOrWhiteSpace(labSheetModelAndA1Sheet.LabSheetA1Sheet.SalinitiesReadMonth) ? 1 : int.Parse(labSheetModelAndA1Sheet.LabSheetA1Sheet.SalinitiesReadMonth));
                int SalinitiesReadDay = (string.IsNullOrWhiteSpace(labSheetModelAndA1Sheet.LabSheetA1Sheet.SalinitiesReadDay) ? 1 : int.Parse(labSheetModelAndA1Sheet.LabSheetA1Sheet.SalinitiesReadDay));
                if (SalinitiesReadYear != 1900)
                {
                    SalinitiesReadDate = new DateTime(SalinitiesReadYear, SalinitiesReadMonth, SalinitiesReadDay);
                }

                labSheetDetailModelNew.SalinitiesReadDate = SalinitiesReadDate;
                labSheetDetailModelNew.ResultsReadBy = labSheetModelAndA1Sheet.LabSheetA1Sheet.ResultsReadBy;

                // ResultsReadDate
                DateTime ResultsReadDate = new DateTime();
                int ResultsReadYear = (string.IsNullOrWhiteSpace(labSheetModelAndA1Sheet.LabSheetA1Sheet.ResultsReadYear) ? 1900 : int.Parse(labSheetModelAndA1Sheet.LabSheetA1Sheet.ResultsReadYear));
                int ResultsReadMonth = (string.IsNullOrWhiteSpace(labSheetModelAndA1Sheet.LabSheetA1Sheet.ResultsReadMonth) ? 1 : int.Parse(labSheetModelAndA1Sheet.LabSheetA1Sheet.ResultsReadMonth));
                int ResultsReadDay = (string.IsNullOrWhiteSpace(labSheetModelAndA1Sheet.LabSheetA1Sheet.ResultsReadDay) ? 1 : int.Parse(labSheetModelAndA1Sheet.LabSheetA1Sheet.ResultsReadDay));
                if (ResultsReadYear != 1900)
                {
                    ResultsReadDate = new DateTime(ResultsReadYear, ResultsReadMonth, ResultsReadDay);
                }

                labSheetDetailModelNew.ResultsReadDate = ResultsReadDate;
                labSheetDetailModelNew.ResultsRecordedBy = labSheetModelAndA1Sheet.LabSheetA1Sheet.ResultsRecordedBy;

                // ResultsRecordedDate
                DateTime ResultsRecordedDate = new DateTime();
                int ResultsRecordedYear = (string.IsNullOrWhiteSpace(labSheetModelAndA1Sheet.LabSheetA1Sheet.ResultsRecordedYear) ? 1900 : int.Parse(labSheetModelAndA1Sheet.LabSheetA1Sheet.ResultsRecordedYear));
                int ResultsRecordedMonth = (string.IsNullOrWhiteSpace(labSheetModelAndA1Sheet.LabSheetA1Sheet.ResultsRecordedMonth) ? 1 : int.Parse(labSheetModelAndA1Sheet.LabSheetA1Sheet.ResultsRecordedMonth));
                int ResultsRecordedDay = (string.IsNullOrWhiteSpace(labSheetModelAndA1Sheet.LabSheetA1Sheet.ResultsRecordedDay) ? 1 : int.Parse(labSheetModelAndA1Sheet.LabSheetA1Sheet.ResultsRecordedDay));
                if (ResultsRecordedYear != 1900)
                {
                    ResultsRecordedDate = new DateTime(ResultsRecordedYear, ResultsRecordedMonth, ResultsRecordedDay);
                }

                labSheetDetailModelNew.ResultsRecordedDate = ResultsRecordedDate;

                labSheetDetailModelNew.DailyDuplicateRLog = 0.0f;
                if (!string.IsNullOrWhiteSpace(labSheetModelAndA1Sheet.LabSheetA1Sheet.DailyDuplicateRLog) && !labSheetModelAndA1Sheet.LabSheetA1Sheet.DailyDuplicateRLog.StartsWith("N"))
                {
                    try
                    {
                        labSheetDetailModelNew.DailyDuplicateRLog = float.Parse(labSheetModelAndA1Sheet.LabSheetA1Sheet.DailyDuplicateRLog);
                    }
                    catch (Exception)
                    {
                        // nothing
                    }
                }

                labSheetDetailModelNew.DailyDuplicatePrecisionCriteria = null;
                labSheetDetailModelNew.DailyDuplicateRLog = null;
                labSheetDetailModelNew.DailyDuplicateAcceptable = null;
                labSheetDetailModelNew.IntertechDuplicatePrecisionCriteria = null;
                labSheetDetailModelNew.IntertechDuplicateRLog = null;
                labSheetDetailModelNew.IntertechDuplicateAcceptable = null;
                labSheetDetailModelNew.IntertechReadAcceptable = null;

                if (!string.IsNullOrWhiteSpace(labSheetModelAndA1Sheet.LabSheetA1Sheet.DailyDuplicatePrecisionCriteria))
                {
                    labSheetDetailModelNew.DailyDuplicatePrecisionCriteria = float.Parse(labSheetModelAndA1Sheet.LabSheetA1Sheet.DailyDuplicatePrecisionCriteria);
                }

                if (!string.IsNullOrWhiteSpace(labSheetModelAndA1Sheet.LabSheetA1Sheet.DailyDuplicateRLog))
                {
                    if (!labSheetModelAndA1Sheet.LabSheetA1Sheet.DailyDuplicateRLog.StartsWith("N"))
                    {
                        labSheetDetailModelNew.DailyDuplicateRLog = float.Parse(labSheetModelAndA1Sheet.LabSheetA1Sheet.DailyDuplicateRLog);
                    }
                }

                if (!string.IsNullOrWhiteSpace(labSheetModelAndA1Sheet.LabSheetA1Sheet.DailyDuplicateAcceptableOrUnacceptable))
                {
                    labSheetDetailModelNew.DailyDuplicateAcceptable = (labSheetModelAndA1Sheet.LabSheetA1Sheet.DailyDuplicateAcceptableOrUnacceptable != "Acceptable" ? false : true);
                }

                if (!string.IsNullOrWhiteSpace(labSheetModelAndA1Sheet.LabSheetA1Sheet.IntertechDuplicatePrecisionCriteria))
                {
                    labSheetDetailModelNew.IntertechDuplicatePrecisionCriteria = float.Parse(labSheetModelAndA1Sheet.LabSheetA1Sheet.IntertechDuplicatePrecisionCriteria);
                }

                if (!string.IsNullOrWhiteSpace(labSheetModelAndA1Sheet.LabSheetA1Sheet.IntertechDuplicateRLog))
                {
                    if (!labSheetModelAndA1Sheet.LabSheetA1Sheet.IntertechDuplicateRLog.StartsWith("N"))
                    {
                        labSheetDetailModelNew.IntertechDuplicateRLog = float.Parse(labSheetModelAndA1Sheet.LabSheetA1Sheet.IntertechDuplicateRLog);
                    }
                }

                if (!string.IsNullOrWhiteSpace(labSheetModelAndA1Sheet.LabSheetA1Sheet.IntertechDuplicateAcceptableOrUnacceptable))
                {
                    labSheetDetailModelNew.IntertechDuplicateAcceptable = (labSheetModelAndA1Sheet.LabSheetA1Sheet.IntertechDuplicateAcceptableOrUnacceptable != "Acceptable" ? false : true);
                }

                if (!string.IsNullOrWhiteSpace(labSheetModelAndA1Sheet.LabSheetA1Sheet.IntertechReadAcceptableOrUnacceptable))
                {
                    labSheetDetailModelNew.IntertechReadAcceptable = (labSheetModelAndA1Sheet.LabSheetA1Sheet.IntertechReadAcceptableOrUnacceptable != "Acceptable" ? false : true);
                }
            }

            LabSheetDetailService labSheetDetailService = new LabSheetDetailService(LanguageEnum.en, _User);

            LabSheetDetailModel labSheetDetailModelExist = labSheetDetailService.GetLabSheetDetailModelExistDB(labSheetDetailModelNew);
            if (!string.IsNullOrWhiteSpace(labSheetDetailModelExist.Error))
            {
                labSheetDetailModelExist = labSheetDetailService.PostAddLabSheetDetailDB(labSheetDetailModelNew);
                if (!string.IsNullOrWhiteSpace(labSheetDetailModelExist.Error))
                {
                    richTextBoxStatus.AppendText("Lab sheet detail could not be loaded to the local DB. Error [" + labSheetDetailModelExist.Error + "]\r\n");
                    return labSheetDetailModelExist.Error;
                }
            }
            else
            {
                labSheetDetailModelNew.LabSheetDetailID = labSheetDetailModelExist.LabSheetDetailID;
                labSheetDetailModelExist = labSheetDetailService.PostUpdateLabSheetDetailDB(labSheetDetailModelNew);
                if (!string.IsNullOrWhiteSpace(labSheetDetailModelExist.Error))
                {
                    richTextBoxStatus.AppendText("Lab sheet detail could not be loaded to the local DB. Error [" + labSheetDetailModelExist.Error + "]\r\n");
                    return labSheetDetailModelExist.Error;
                }
            }

            retStr = UploadLabSheetTubeMPNDetailInDB(labSheetDetailModelExist.LabSheetDetailID, labSheetModelAndA1Sheet.LabSheetA1Sheet.LabSheetA1MeasurementList);
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
                LabSheetTubeMPNDetailModel labSheetTubeMPNDetailModelNew = new LabSheetTubeMPNDetailModel();
                labSheetTubeMPNDetailModelNew.LabSheetDetailID = LabSheetDetailID;
                labSheetTubeMPNDetailModelNew.Ordinal = Ordinal;
                labSheetTubeMPNDetailModelNew.MWQMSiteTVItemID = labSheetA1Measurement.TVItemID;
                labSheetTubeMPNDetailModelNew.SampleDateTime = labSheetA1Measurement.Time;
                labSheetTubeMPNDetailModelNew.MPN = labSheetA1Measurement.MPN;
                labSheetTubeMPNDetailModelNew.Tube10 = labSheetA1Measurement.Tube10;
                labSheetTubeMPNDetailModelNew.Tube1_0 = labSheetA1Measurement.Tube1_0;
                labSheetTubeMPNDetailModelNew.Tube0_1 = labSheetA1Measurement.Tube0_1;
                labSheetTubeMPNDetailModelNew.Salinity = labSheetA1Measurement.Salinity;
                labSheetTubeMPNDetailModelNew.Temperature = labSheetA1Measurement.Temperature;
                labSheetTubeMPNDetailModelNew.ProcessedBy = labSheetA1Measurement.ProcessedBy;
                labSheetTubeMPNDetailModelNew.SampleType = (SampleTypeEnum)labSheetA1Measurement.SampleType;
                labSheetTubeMPNDetailModelNew.SiteComment = labSheetA1Measurement.SiteComment;

                LabSheetTubeMPNDetailModel labSheetTubeMPNDetailModelExist = labSheetTubeMPNDetailService.GetLabSheetTubeMPNDetailModelExistDB(labSheetTubeMPNDetailModelNew);
                if (!string.IsNullOrWhiteSpace(labSheetTubeMPNDetailModelExist.Error))
                {
                    labSheetTubeMPNDetailModelExist = labSheetTubeMPNDetailService.PostAddLabSheetTubeMPNDetailDB(labSheetTubeMPNDetailModelNew);
                    if (!string.IsNullOrWhiteSpace(labSheetTubeMPNDetailModelExist.Error))
                        return labSheetTubeMPNDetailModelExist.Error;
                }
                else
                {
                    labSheetTubeMPNDetailModelNew.LabSheetTubeMPNDetailID = labSheetTubeMPNDetailModelExist.LabSheetTubeMPNDetailID;
                    labSheetTubeMPNDetailModelExist = labSheetTubeMPNDetailService.PostUpdateLabSheetTubeMPNDetailDB(labSheetTubeMPNDetailModelNew);
                    if (!string.IsNullOrWhiteSpace(labSheetTubeMPNDetailModelExist.Error))
                        return labSheetTubeMPNDetailModelExist.Error;
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

                        TVItemModel tvItemModelSubsector = new TVItemModel();
                        TVItemModel tvItemModelCountry = new TVItemModel();
                        TVItemModel tvItemModelProvince = new TVItemModel();
                        LabSheetModelAndA1Sheet labSheetModelAndA1Sheet = new LabSheetModelAndA1Sheet();
                        using (TransactionScope ts = new TransactionScope())
                        {
                            LabSheetService labSheetService = new LabSheetService(LanguageEnum.en, _User);

                            LabSheetModel labSheetModelRet = labSheetService.AddOrUpdateLabSheetDB(FullLabSheetText);
                            if (!string.IsNullOrWhiteSpace(labSheetModelRet.Error))
                            {
                                richTextBoxStatus.AppendText("Lab sheet adding error OtherServerLabSheetID [" + OtherServerLabSheetID.ToString() + "]" + labSheetModelRet.Error + "]\r\n");
                                return;
                            }

                            TVItemService tvItemService = new TVItemService(LanguageEnum.en, _User);

                            tvItemModelSubsector = tvItemService.GetTVItemModelWithTVItemIDDB(SubsectorTVItemID);
                            if (!string.IsNullOrWhiteSpace(tvItemModelSubsector.Error))
                            {
                                richTextBoxStatus.AppendText("Lab sheet parsing error OtherServerLabSheetID [" + OtherServerLabSheetID.ToString() + "]" + tvItemModelSubsector.Error + "]\r\n");
                                return;
                            }

                            List<TVItemModel> tvItemModelList = tvItemService.GetParentsTVItemModelList(tvItemModelSubsector.TVPath);
                            foreach (TVItemModel tvItemModel in tvItemModelList)
                            {
                                if (tvItemModel.TVType == TVTypeEnum.Province)
                                {
                                    tvItemModelProvince = tvItemModel;
                                }
                                if (tvItemModel.TVType == TVTypeEnum.Country)
                                {
                                    tvItemModelCountry = tvItemModel;
                                }
                            }

                            labSheetModelAndA1Sheet.LabSheetModel = labSheetModelRet;
                            labSheetModelAndA1Sheet.LabSheetA1Sheet = labSheetService.ParseLabSheetA1WithLabSheetID(labSheetModelRet.LabSheetID);
                            if (!string.IsNullOrWhiteSpace(labSheetModelAndA1Sheet.LabSheetA1Sheet.Error))
                            {
                                richTextBoxStatus.AppendText("Lab sheet parsing error OtherServerLabSheetID [" + OtherServerLabSheetID.ToString() + "]" + labSheetModelAndA1Sheet.LabSheetA1Sheet.Error + "]\r\n");
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

                            string retStr2 = UploadLabSheetDetailInDB(labSheetModelAndA1Sheet);
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

                        string href = "http://wmon01dtchlebl2/csspwebtools/en-CA/#!View/" + (tvItemModelCountry.TVText + "-" + tvItemModelProvince.TVText).Replace(" ", "-") + "|||" + tvItemModelProvince.TVItemID.ToString() + "|||010003030200000000000000000000";

                        if (labSheetModelAndA1Sheet.LabSheetA1Sheet.LabSheetA1MeasurementList.Where(c => c.MPN != null && c.MPN >= MPNLimitForEmail).Any())
                        {
                            SendNewLabSheetEmailBigMPN(href, tvItemModelProvince, tvItemModelSubsector, labSheetModelAndA1Sheet);
                        }
                        else
                        {
                            SendNewLabSheetEmail(href, tvItemModelProvince, tvItemModelSubsector, labSheetModelAndA1Sheet);
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
