using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Security.Principal;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using CSSPTaskRunnerHelper;

namespace CSSPTaskRunner
{
    public partial class CSSPTaskRunner : Form
    {
        #region Variables
        #endregion Variables

        #region Properties
        public TaskRunnerHelper taskRunnerHelper { get; set; }
        #endregion Properties

        #region Constructors
        public CSSPTaskRunner()
        {
            InitializeComponent();
            Setup();
            taskRunnerHelper = new TaskRunnerHelper();
            taskRunnerHelper.MessageHandler += TaskRunnerHelper_MessageHandler; 
        }

        #endregion Constructors

        #region Events
        private void butGetTasksInformation_Click(object sender, EventArgs e)
        {
            taskRunnerHelper.ShowTasksInformation();
        }
        private void TaskRunnerHelper_MessageHandler(object sender, TaskRunnerHelperBase.MessageEventArgs e)
        {
            TaskRunnerHelperBase.MessageEventArgs message = e;
            switch (e.Type)
            {
                case TaskRunnerHelperBase.MessageTypeEnum.Error:
                    {
                        richTextBoxStatus.AppendText($"\r\nERROR: { e.Message }\r\n\r\n");
                        richTextBoxStatus.Refresh();
                        Application.DoEvents();
                    }
                    break;
                case TaskRunnerHelperBase.MessageTypeEnum.Status:
                    {
                        lblStatus.Text = e.Message;
                        lblStatus.Refresh();
                        Application.DoEvents();
                    }
                    break;
                case TaskRunnerHelperBase.MessageTypeEnum.Permanent:
                    {
                        richTextBoxStatus.AppendText($"{ e.Message }\r\n");
                        richTextBoxStatus.Refresh();
                        Application.DoEvents();
                    }
                    break;
                case TaskRunnerHelperBase.MessageTypeEnum.Clear:
                    {
                        richTextBoxStatus.Text = "";
                        richTextBoxStatus.Refresh();
                        Application.DoEvents();
                    }
                    break;
                default:
                    break;
            }
        }
        private void timerCheckTask_Tick(object sender, EventArgs e)
        {
            DoTimerCheckTask();
        }
        #endregion Events

        #region Functions
        private void Setup()
        {
            lblStatus.Text = "";
            richTextBoxStatus.Text = "";
            timerCheckTask.Enabled = true;
        }
        public void DoTimerCheckTask()
        {
            timerCheckTask.Enabled = false;
            taskRunnerHelper.GetNextTask();
            timerCheckTask.Enabled = true;
        }

        public class RunDateColNumber
        {
            public RunDateColNumber()
            {

            }
            public int MWQMRunTVItemID { get; set; }
            public DateTime RunDate { get; set; }
            public int ColNumber { get; set; }
            public bool Used { get; set; }
        }
        public class StatRunSite
        {
            public StatRunSite()
            {

            }

            public int? SampleCount { get; set; }
            public int? PeriodStart { get; set; }
            public int? PeriodEnd { get; set; }
            public int? MinFC { get; set; }
            public int? MaxFC { get; set; }
            public double? GM { get; set; }
            public double? Med { get; set; }
            public double? P90 { get; set; }
            public double? P43 { get; set; }
            public double? P260 { get; set; }
            public string Letter { get; set; }
            public int Color { get; set; }

        }
        public class SiteRowNumber
        {
            public SiteRowNumber()
            {

            }
            public int MWQMSiteTVItemID { get; set; }
            public string SiteName { get; set; }
            public int RowNumber { get; set; }
        }
        public class RowAndType
        {
            public RowAndType()
            {

            }
            public int RowNumber { get; set; }
            public ExcelExportShowDataTypeEnum ExcelExportShowDataType { get; set; }
        }
        private string GetTideInitial(TideTextEnum? tideText)
        {
            if (tideText == null)
            {
                return "--";
            }

            switch (tideText)
            {
                case TideTextEnum.LowTide:
                    return "LT";
                case TideTextEnum.LowTideFalling:
                    return "LF";
                case TideTextEnum.LowTideRising:
                    return "LR";
                case TideTextEnum.MidTide:
                    return "MT";
                case TideTextEnum.MidTideFalling:
                    return "MF";
                case TideTextEnum.MidTideRising:
                    return "MR";
                case TideTextEnum.HighTide:
                    return "HT";
                case TideTextEnum.HighTideFalling:
                    return "HF";
                case TideTextEnum.HighTideRising:
                    return "HR";
                default:
                    return "--";
            }
        }


        private int GetLastClassificationColor(MWQMSiteLatestClassificationEnum? mwqmSiteLatestClassification)
        {
            if (mwqmSiteLatestClassification == null)
            {
                return 16777215;
            }

            switch (mwqmSiteLatestClassification)
            {
                case MWQMSiteLatestClassificationEnum.Approved:
                    return 5287936;
                case MWQMSiteLatestClassificationEnum.ConditionallyApproved:
                    return 5287936;
                case MWQMSiteLatestClassificationEnum.ConditionallyRestricted:
                    return 0;
                case MWQMSiteLatestClassificationEnum.Prohibited:
                    return 0;
                case MWQMSiteLatestClassificationEnum.Restricted:
                    return 255;
                case MWQMSiteLatestClassificationEnum.Unclassified:
                    return 16777215;
                default:
                    return 16777215;
            }
        }

        private string GetLastClassificationInitial(MWQMSiteLatestClassificationEnum? mwqmSiteLatestClassification)
        {
            if (mwqmSiteLatestClassification == null)
            {
                return "";
            }

            switch (mwqmSiteLatestClassification)
            {
                case MWQMSiteLatestClassificationEnum.Approved:
                    return (LanguageRequest == LanguageEnum.fr ? "A" : "A");
                case MWQMSiteLatestClassificationEnum.ConditionallyApproved:
                    return (LanguageRequest == LanguageEnum.fr ? "CA" : "AC");
                case MWQMSiteLatestClassificationEnum.ConditionallyRestricted:
                    return (LanguageRequest == LanguageEnum.fr ? "CR" : "RC");
                case MWQMSiteLatestClassificationEnum.Prohibited:
                    return (LanguageRequest == LanguageEnum.fr ? "P" : "P");
                case MWQMSiteLatestClassificationEnum.Restricted:
                    return (LanguageRequest == LanguageEnum.fr ? "R" : "R");
                case MWQMSiteLatestClassificationEnum.Unclassified:
                    return (LanguageRequest == LanguageEnum.fr ? "" : "");
                default:
                    return "";
            }
        }

        private void SetupParametersAndBasicTextOnSheet1(Microsoft.Office.Interop.Excel.Application xlApp, Microsoft.Office.Interop.Excel.Workbook wb, MWQMAnalysisReportParameterModel mwqmAnalysisReportParameterModel)
        {
            _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, 7);

            CSSPEnumsDLL.Services.BaseEnumService _BaseEnumService = new CSSPEnumsDLL.Services.BaseEnumService(LanguageEnum.en);
            Microsoft.Office.Interop.Excel.Worksheet ws = (Microsoft.Office.Interop.Excel.Worksheet)wb.Worksheets[1];
            Microsoft.Office.Interop.Excel.Range range = ws.get_Range("A1:A1");
            if (ws == null)
            {
                Console.WriteLine("Worksheet could not be created. Check that your office installation and project references are correct.");
            }
            ws.Activate();
            ws.Name = "Stat and Data";
            range = xlApp.get_Range("A1:A1");
            range.Value = "Parameters";

            range = xlApp.get_Range("A1:J1");
            range.Select();
            range.HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
            range.VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
            range.Merge();
            xlApp.Selection.Borders().LineStyle = Microsoft.Office.Interop.Excel.Constants.xlNone;
            xlApp.Selection.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            xlApp.Selection.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;

            List<string> textList = new List<string>() { "", "Run Date\n\nRain Day", "Run Day (0)", "0-24h (-1)", "24-48h (-2)", "48-72h (-3)",
                "72-96h (-4)", "(-5)", "(-6)", "(-7)", "(-8)", "(-9)", "(-10)", "Start Tide", "End Tide" };

            for (int i = 1; i < 15; i++)
            {
                range = xlApp.get_Range("K" + i + ":L" + i);
                range.Select();
                range.HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlRight;
                range.VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
                range.Merge();

                xlApp.Selection.Borders().LineStyle = Microsoft.Office.Interop.Excel.Constants.xlNone;
                xlApp.Selection.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                xlApp.Selection.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft).Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                range.Value = textList[i];
            }

            ws.Columns["A:A"].ColumnWidth = 4.89;
            ws.Columns["B:B"].ColumnWidth = 2.11;
            ws.Columns["C:C"].ColumnWidth = 6.33;
            ws.Columns["D:D"].ColumnWidth = 7.33;
            ws.Columns["E:E"].ColumnWidth = 5.22;
            ws.Columns["F:F"].ColumnWidth = 5.44;
            ws.Columns["G:G"].ColumnWidth = 5.22;
            ws.Columns["H:H"].ColumnWidth = 4.78;
            ws.Columns["I:I"].ColumnWidth = 3.89;
            ws.Columns["J:J"].ColumnWidth = 5.33;
            ws.Columns["K:K"].ColumnWidth = 5.67;
            ws.Columns["L:L"].ColumnWidth = 1.22;
            ws.Rows["1:1"].RowHeight = 43;

            textList = new List<string>() { "", "", "Between", "And", "Select Full Year", "Runs", "Sal", "Short Range", "Mid Range", "Calculation" };
            for (int i = 2; i < 10; i++)
            {
                range = xlApp.get_Range("D" + i + ":E" + i);
                range.Select();
                range.HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlRight;
                range.VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
                range.Merge();
                range.Value = textList[i];
            }

            textList = new List<string>() { "", "", "'" + mwqmAnalysisReportParameterModel.StartDate.ToString("yyyy MMM dd"),
                "'" + mwqmAnalysisReportParameterModel.EndDate.ToString("yyyy MMM dd"),
                "'" + mwqmAnalysisReportParameterModel.FullYear.ToString(),
                "'" + mwqmAnalysisReportParameterModel.NumberOfRuns.ToString(),
                "'" + mwqmAnalysisReportParameterModel.SalinityHighlightDeviationFromAverage.ToString(),
                "'" + mwqmAnalysisReportParameterModel.ShortRangeNumberOfDays.ToString(),
                "'" + mwqmAnalysisReportParameterModel.MidRangeNumberOfDays.ToString(),
                "'" + _BaseEnumService.GetEnumText_AnalysisCalculationTypeEnum(mwqmAnalysisReportParameterModel.AnalysisCalculationType) };
            for (int i = 2; i < 10; i++)
            {
                range = xlApp.get_Range("F" + i + ":G" + i);
                range.Select();
                range.HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlRight;
                range.VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
                range.Merge();
                range.Value = textList[i];
            }

            range = xlApp.get_Range("D2:G9");
            range.Select();
            xlApp.Selection.Borders().LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            xlApp.Selection.Borders().Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;

            for (int i = 12; i < 14; i++)
            {
                range = xlApp.get_Range("C" + i + ":C" + i);
                range.Select();
                range.HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlRight;
                range.VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
                range.Merge();
                range.Value = (i == 12 ? "Dry" : "Wet");
            }

            for (int i = 11; i < 14; i++)
            {
                range = xlApp.get_Range("D" + i + ":D" + i);
                range.Select();
                range.HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
                range.VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
                range.Merge();
                range.Value = (i == 11 ? "0-24h" : (i == 12 ? mwqmAnalysisReportParameterModel.DryLimit24h.ToString() : mwqmAnalysisReportParameterModel.WetLimit24h.ToString()));
            }

            for (int i = 11; i < 14; i++)
            {
                range = xlApp.get_Range("E" + i + ":E" + i);
                range.Select();
                range.HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
                range.VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
                range.Merge();
                range.Value = (i == 11 ? "0-48h" : (i == 12 ? mwqmAnalysisReportParameterModel.DryLimit48h.ToString() : mwqmAnalysisReportParameterModel.WetLimit48h.ToString()));
            }

            for (int i = 11; i < 14; i++)
            {
                range = xlApp.get_Range("F" + i + ":F" + i);
                range.Select();
                range.HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
                range.VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
                range.Merge();
                range.Value = (i == 11 ? "0-72h" : (i == 12 ? mwqmAnalysisReportParameterModel.DryLimit72h.ToString() : mwqmAnalysisReportParameterModel.WetLimit72h.ToString()));
            }

            for (int i = 11; i < 14; i++)
            {
                range = xlApp.get_Range("G" + i + ":G" + i);
                range.Select();
                range.HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
                range.VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
                range.Merge();
                range.Value = (i == 11 ? "0-96h" : (i == 12 ? mwqmAnalysisReportParameterModel.DryLimit96h.ToString() : mwqmAnalysisReportParameterModel.WetLimit96h.ToString()));
            }

            range = xlApp.get_Range("D11:G11");
            range.Select();
            xlApp.Selection.Borders().LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            xlApp.Selection.Borders().Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
            xlApp.Selection.Font.Bold = true;

            range = xlApp.get_Range("C12:G13");
            range.Select();
            xlApp.Selection.Borders().LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            xlApp.Selection.Borders().Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;

            range = xlApp.get_Range("C12:C13");
            range.Select();
            xlApp.Selection.Font.Bold = true;

            textList = new List<string>() { "Site", "Samples", "Period", "Min FC", "Max FC", "GMean", "Median", "P90", "% > 43", "% > 260" };
            List<string> LetterList = new List<string>() { "A", "C", "D", "E", "F", "G", "H", "I", "J", "K" };
            for (int i = 0; i < 10; i++)
            {
                range = xlApp.get_Range(LetterList[i] + "15:" + LetterList[i] + "15");
                range.Select();
                range.Value = textList[i];
            }

            range = xlApp.get_Range("A15:K15");
            range.Select();
            xlApp.Selection.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            xlApp.Selection.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;

            range = xlApp.get_Range("M15:M15");
            range.Select();
            List<string> showDataTypeTextList = mwqmAnalysisReportParameterModel.ShowDataTypes.Split(",".ToCharArray(), StringSplitOptions.RemoveEmptyEntries).ToList();
            string M15Text = "      ";
            foreach (string s in showDataTypeTextList)
            {
                M15Text = M15Text + _BaseEnumService.GetEnumText_ExcelExportShowDataTypeEnum(((ExcelExportShowDataTypeEnum)int.Parse(s))) + ", ";
            }
            range.Value = M15Text;
            xlApp.Selection.WrapText = false;
            range.HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlLeft;

            ws.Cells.Select();
            xlApp.Selection.Font.Size = 10;

            range = xlApp.get_Range("A1:A1");
            range.Select();

            _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, 10);

        }
        private void SetupParametersAndBasicTextOnSheet2(Microsoft.Office.Interop.Excel.Application xlApp, Microsoft.Office.Interop.Excel.Workbook wb, MWQMAnalysisReportParameterModel mwqmAnalysisReportParameterModel)
        {
            _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, 3);

            CSSPEnumsDLL.Services.BaseEnumService _BaseEnumService = new CSSPEnumsDLL.Services.BaseEnumService(LanguageEnum.en);
            Microsoft.Office.Interop.Excel.Worksheet ws = (Microsoft.Office.Interop.Excel.Worksheet)wb.Worksheets[2];
            Microsoft.Office.Interop.Excel.Range range = ws.get_Range("A1:A1");
            if (ws == null)
            {
                Console.WriteLine("Worksheet could not be created. Check that your office installation and project references are correct.");
            }

            ws.Activate();
            ws.Name = "Help";
            range = xlApp.get_Range("A1:G1");
            range.Select();
            range.HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
            range.VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
            range.Merge();
            xlApp.Selection.Borders().LineStyle = Microsoft.Office.Interop.Excel.Constants.xlNone;
            xlApp.Selection.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            xlApp.Selection.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
            range.Value = "Color and letter schema";

            List<string> LetterList = new List<string>()
            {
                "F","E","D","C","B","A","F","E","D","C","B","A","F","E","D","C","B","A",
            };
            List<string> RangeList = new List<string>()
            {
               "GM > 181.33 or Med > 181.33 or P90 > 460.0 or % > 260 > 18.33",
               "GM > 162.67 or Med > 162.67 or P90 > 420.0 or % > 260 > 16.67",
               "GM > 144.0 or Med > 144.0 or P90 > 380.0 or % > 260 > 15.0",
               "GM > 125.33 or Med > 125.33 or P90 > 340.0 or % > 260 > 13.33",
               "GM > 106.67 or Med > 106.67 or P90 > 300.0 or % > 260 > 11.67",
               "GM > 88 or Med > 88 or P90 > 260 or % > 260 > 10",
               "GM > 75.67 or Med > 75.67 or P90 > 223.83 or % > 43 > 26.67",
               "GM > 63.33 or Med > 63.33 or P90 > 187.67 or % > 43 > 23.33",
               "GM > 51.0 or Med > 51.0 or P90 > 151.5 or % > 43 > 20.0",
               "GM > 38.67 or Med > 38.67 or P90 > 115.33 or % > 43 > 16.67",
               "GM > 26.33 or Med > 26.33 or P90 > 79.17 or % > 43 > 13.33",
               "GM > 14 or Med > 14 or P90 > 43 or % > 43 > 10",
               "GM > 11.67 or Med > 11.67 or P90 > 35.83 or % > 43 > 8.33",
               "GM > 9.33 or Med > 9.33 or P90 > 28.67 or % > 43 > 6.67",
               "GM > 7.0 or Med > 7.0 or P90 > 21.5 or % > 43 > 5.0",
               "GM > 4.67 or Med > 4.67 or P90 > 14.33 or % > 43 > 3.33",
               "GM > 2.33 or Med > 2.33 or P90 > 7.17 or % > 43 > 1.67",
               "Everything else",
            };

            List<string> BGColorList = new List<string>
            {
                "16746632",
                "16751001",
                "16755370",
                "16759739",
                "16764108",
                "16768477",
                "170",
                "204",
                "1118718",
                "4474111",
                "10066431",
                "13421823",
                "13434828",
                "10092441",
                "4521796",
                "1179409",
                "47872",
                "39168",
            };

            for (int i = 0, count = LetterList.Count; i < count; i++)
            {
                range = xlApp.get_Range("A" + (i + 3).ToString() + ":A" + (i + 3).ToString());
                range.Select();
                range.HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
                range.VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
                xlApp.Selection.Borders().LineStyle = Microsoft.Office.Interop.Excel.Constants.xlNone;
                xlApp.Selection.Borders().LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                xlApp.Selection.Borders().Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                range.Value = LetterList[i];

                xlApp.Selection.Interior.Color = int.Parse(BGColorList[i]);

                range = xlApp.get_Range("B" + (i + 3).ToString() + ":B" + (i + 3).ToString());
                range.Select();
                range.Value = RangeList[i];
            }

            ws.Columns["A:A"].ColumnWidth = 2.11;

            _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, 5);

        }
        private void SetupStatOnSheet1(Microsoft.Office.Interop.Excel.Application xlApp, Microsoft.Office.Interop.Excel.Workbook wb, MWQMAnalysisReportParameterModel mwqmAnalysisReportParameterModel)
        {
            int LatestYear = 0;
            List<int> RunUsedColNumberList = new List<int>();
            TVItemService tvItemService = new TVItemService(LanguageRequest, _User);
            List<RunDateColNumber> runDateColNumberList = new List<RunDateColNumber>();
            List<RowAndType> rowAndTypeList = new List<RowAndType>();

            List<SiteRowNumber> siteRowNumberList = new List<SiteRowNumber>();
            List<ExcelExportShowDataTypeEnum> showDataTypeList = new List<ExcelExportShowDataTypeEnum>();
            List<int> MWQMRunTVItemIDToOmitList = new List<int>();

            string[] showDataTypeTextList = mwqmAnalysisReportParameterModel.ShowDataTypes.Split(",".ToCharArray(), StringSplitOptions.RemoveEmptyEntries);
            foreach (string s in showDataTypeTextList)
            {
                showDataTypeList.Add((ExcelExportShowDataTypeEnum)int.Parse(s));
            }

            string[] runDateTextList = mwqmAnalysisReportParameterModel.RunsToOmit.Split(",".ToCharArray(), StringSplitOptions.RemoveEmptyEntries);
            foreach (string s in runDateTextList)
            {
                MWQMRunTVItemIDToOmitList.Add(int.Parse(s));
            }

            CSSPEnumsDLL.Services.BaseEnumService _BaseEnumService = new CSSPEnumsDLL.Services.BaseEnumService(LanguageEnum.en);
            Microsoft.Office.Interop.Excel.Worksheet ws = (Microsoft.Office.Interop.Excel.Worksheet)wb.Worksheets[1];
            Microsoft.Office.Interop.Excel.Range range = ws.get_Range("A1:A1");
            if (ws == null)
            {
                Console.WriteLine("Worksheet could not be created. Check that your office installation and project references are correct.");
            }
            ws.Activate();

            MWQMSubsectorService _MWQMSubsectorService = new MWQMSubsectorService(LanguageEnum.en, _User);
            MWQMSubsectorAnalysisModel mwqmSubsectorAnalysisModel = _MWQMSubsectorService.GetMWQMSubsectorAnalysisModel(mwqmAnalysisReportParameterModel.SubsectorTVItemID);

            foreach (MWQMSampleAnalysisModel mwqmSampleAnalysisModel in mwqmSubsectorAnalysisModel.MWQMSampleAnalysisModelList)
            {
                mwqmSampleAnalysisModel.SampleDateTime_Local = new DateTime(mwqmSampleAnalysisModel.SampleDateTime_Local.Year, mwqmSampleAnalysisModel.SampleDateTime_Local.Month, mwqmSampleAnalysisModel.SampleDateTime_Local.Day);
            }

            foreach (MWQMRunAnalysisModel mwqmRunAnalysisModel in mwqmSubsectorAnalysisModel.MWQMRunAnalysisModelList)
            {
                mwqmRunAnalysisModel.DateTime_Local = new DateTime(mwqmRunAnalysisModel.DateTime_Local.Year, mwqmRunAnalysisModel.DateTime_Local.Month, mwqmRunAnalysisModel.DateTime_Local.Day);
            }

            _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, 20);

            int CountRun = mwqmSubsectorAnalysisModel.MWQMRunAnalysisModelList.Count();
            for (int i = 0, count = mwqmSubsectorAnalysisModel.MWQMRunAnalysisModelList.Count(); i < count; i++)
            {
                if (i % 20 == 0)
                {
                    int Percent = (int)(20.0D + (30.0D * ((double)i / (double)CountRun)));
                    _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, Percent);
                }

                if (i == 0)
                {
                    LatestYear = mwqmSubsectorAnalysisModel.MWQMRunAnalysisModelList[i].DateTime_Local.Year;
                }
                ws.Cells[1, 13 + i] = mwqmSubsectorAnalysisModel.MWQMRunAnalysisModelList[i].DateTime_Local.ToString("yyyy\nMMM\ndd");
                ws.Cells[2, 13 + i] = (mwqmSubsectorAnalysisModel.MWQMRunAnalysisModelList[i].RainDay0_mm == null
                    ? "--" : Convert.ToInt32(mwqmSubsectorAnalysisModel.MWQMRunAnalysisModelList[i].RainDay0_mm).ToString());
                ws.Cells[3, 13 + i] = (mwqmSubsectorAnalysisModel.MWQMRunAnalysisModelList[i].RainDay1_mm == null
                    ? "--" : Convert.ToInt32(mwqmSubsectorAnalysisModel.MWQMRunAnalysisModelList[i].RainDay1_mm).ToString());
                ws.Cells[4, 13 + i] = (mwqmSubsectorAnalysisModel.MWQMRunAnalysisModelList[i].RainDay2_mm == null
                    ? "--" : Convert.ToInt32(mwqmSubsectorAnalysisModel.MWQMRunAnalysisModelList[i].RainDay2_mm).ToString());
                ws.Cells[5, 13 + i] = (mwqmSubsectorAnalysisModel.MWQMRunAnalysisModelList[i].RainDay3_mm == null
                    ? "--" : Convert.ToInt32(mwqmSubsectorAnalysisModel.MWQMRunAnalysisModelList[i].RainDay3_mm).ToString());
                ws.Cells[6, 13 + i] = (mwqmSubsectorAnalysisModel.MWQMRunAnalysisModelList[i].RainDay4_mm == null
                    ? "--" : Convert.ToInt32(mwqmSubsectorAnalysisModel.MWQMRunAnalysisModelList[i].RainDay4_mm).ToString());
                ws.Cells[7, 13 + i] = (mwqmSubsectorAnalysisModel.MWQMRunAnalysisModelList[i].RainDay5_mm == null
                    ? "--" : Convert.ToInt32(mwqmSubsectorAnalysisModel.MWQMRunAnalysisModelList[i].RainDay5_mm).ToString());
                ws.Cells[8, 13 + i] = (mwqmSubsectorAnalysisModel.MWQMRunAnalysisModelList[i].RainDay6_mm == null
                    ? "--" : Convert.ToInt32(mwqmSubsectorAnalysisModel.MWQMRunAnalysisModelList[i].RainDay6_mm).ToString());
                ws.Cells[9, 13 + i] = (mwqmSubsectorAnalysisModel.MWQMRunAnalysisModelList[i].RainDay7_mm == null
                    ? "--" : Convert.ToInt32(mwqmSubsectorAnalysisModel.MWQMRunAnalysisModelList[i].RainDay7_mm).ToString());
                ws.Cells[10, 13 + i] = (mwqmSubsectorAnalysisModel.MWQMRunAnalysisModelList[i].RainDay8_mm == null
                    ? "--" : Convert.ToInt32(mwqmSubsectorAnalysisModel.MWQMRunAnalysisModelList[i].RainDay8_mm).ToString());
                ws.Cells[11, 13 + i] = (mwqmSubsectorAnalysisModel.MWQMRunAnalysisModelList[i].RainDay9_mm == null
                    ? "--" : Convert.ToInt32(mwqmSubsectorAnalysisModel.MWQMRunAnalysisModelList[i].RainDay9_mm).ToString());
                ws.Cells[12, 13 + i] = (mwqmSubsectorAnalysisModel.MWQMRunAnalysisModelList[i].RainDay10_mm == null
                    ? "--" : Convert.ToInt32(mwqmSubsectorAnalysisModel.MWQMRunAnalysisModelList[i].RainDay10_mm).ToString());
                ws.Cells[13, 13 + i] = (mwqmSubsectorAnalysisModel.MWQMRunAnalysisModelList[i].Tide_Start == null
                    ? "--" : GetTideInitial(mwqmSubsectorAnalysisModel.MWQMRunAnalysisModelList[i].Tide_Start));
                ws.Cells[14, 13 + i] = (mwqmSubsectorAnalysisModel.MWQMRunAnalysisModelList[i].Tide_End == null
                    ? "--" : GetTideInitial(mwqmSubsectorAnalysisModel.MWQMRunAnalysisModelList[i].Tide_End));

                ws.Columns[13 + i].ColumnWidth = 4.33;
                range = ws.Columns[13 + i];
                range.Select();
                xlApp.Selection.HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;

                runDateColNumberList.Add(new RunDateColNumber()
                {
                    MWQMRunTVItemID = mwqmSubsectorAnalysisModel.MWQMRunAnalysisModelList[i].MWQMRunTVItemID,
                    RunDate = mwqmSubsectorAnalysisModel.MWQMRunAnalysisModelList[i].DateTime_Local,
                    ColNumber = 13 + i,
                    Used = false,
                });

                if (LatestYear != mwqmSubsectorAnalysisModel.MWQMRunAnalysisModelList[i].DateTime_Local.Year)
                {
                    ws.Columns[13 + i].Select();
                    xlApp.Selection.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft).ColorIndex = 0;
                    xlApp.Selection.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    LatestYear = mwqmSubsectorAnalysisModel.MWQMRunAnalysisModelList[i].DateTime_Local.Year;
                }
            }

            ws.Columns["A:L"].Select();
            xlApp.Selection.HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;

            ws.Range["M15"].Select();
            xlApp.Selection.HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlLeft;

            int RowCount = 16;
            List<MWQMSiteAnalysisModel> mwqmSiteAnalysisModelListAll = mwqmSubsectorAnalysisModel.MWQMSiteAnalysisModelList.Where(c => c.IsActive == true).OrderBy(c => c.MWQMSiteTVText)
                                                                        .Concat(mwqmSubsectorAnalysisModel.MWQMSiteAnalysisModelList.Where(c => c.IsActive == false).OrderBy(c => c.MWQMSiteTVText)).ToList();

            int CountSite = 0;
            int CountSiteTotal = mwqmSiteAnalysisModelListAll.Count();
            foreach (MWQMSiteAnalysisModel mwqmSiteAnalysisModel in mwqmSiteAnalysisModelListAll)
            {
                CountSite += 1;
                if (CountSite % 10 == 0)
                {
                    int Percent = (int)(50.0D + (50.0D * ((double)CountSite / (double)CountSiteTotal)));
                    _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, Percent);
                }

                double? P90 = null;
                double? GeoMean = null;
                double? Median = null;
                double? PercOver43 = null;
                double? PercOver260 = null;

                SiteRowNumber siteRowNumber = new SiteRowNumber() { MWQMSiteTVItemID = mwqmSiteAnalysisModel.MWQMSiteTVItemID, SiteName = mwqmSiteAnalysisModel.MWQMSiteTVText, RowNumber = RowCount };
                siteRowNumberList.Add(siteRowNumber);

                range = ws.Cells[RowCount, 1];
                string classification = GetLastClassificationInitial(mwqmSiteAnalysisModel.MWQMSiteLatestClassification);
                range.Value = "'" + mwqmSiteAnalysisModel.MWQMSiteTVText;
                range.HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
                range.Select();
                if (mwqmSiteAnalysisModel.IsActive == true)
                {
                    xlApp.Selection.Font.Color = 5287936; // green
                }
                else
                {
                    xlApp.Selection.Font.Color = 255; // red
                }

                range = ws.Cells[RowCount, 2];
                range.Value = "'" + (string.IsNullOrWhiteSpace(classification) ? "" : classification);
                range.Select();
                xlApp.Selection.Interior.Color = GetLastClassificationColor(mwqmSiteAnalysisModel.MWQMSiteLatestClassification);
                xlApp.Selection.Font.Color = (classification == "P" ? 16777215 : 0);

                // loading all site sample and doing the stats
                List<MWQMSampleAnalysisModel> mwqmSampleAnalysisForSiteModelList = mwqmSubsectorAnalysisModel.MWQMSampleAnalysisModelList.Where(c => c.MWQMSiteTVItemID == siteRowNumber.MWQMSiteTVItemID).OrderByDescending(c => c.SampleDateTime_Local).ToList();
                List<MWQMSampleAnalysisModel> mwqmSampleAnalysisForSiteModelToUseList = new List<MWQMSampleAnalysisModel>();
                foreach (MWQMSampleAnalysisModel mwqmSampleAnalysisModel in mwqmSampleAnalysisForSiteModelList)
                {
                    if (!MWQMRunTVItemIDToOmitList.Contains(mwqmSampleAnalysisModel.MWQMRunTVItemID))
                    {
                        if ((mwqmSampleAnalysisModel.SampleDateTime_Local <= mwqmAnalysisReportParameterModel.StartDate) && (mwqmSampleAnalysisModel.SampleDateTime_Local >= mwqmAnalysisReportParameterModel.EndDate))
                        {
                            if (mwqmSampleAnalysisForSiteModelToUseList.Count < mwqmAnalysisReportParameterModel.NumberOfRuns)
                            {
                                if (mwqmAnalysisReportParameterModel.AnalysisCalculationType == AnalysisCalculationTypeEnum.WetAllAll)
                                {
                                    MWQMRunAnalysisModel mwqmRunAnalysisModel = (from c in mwqmSubsectorAnalysisModel.MWQMRunAnalysisModelList
                                                                                 where c.DateTime_Local.Year == mwqmSampleAnalysisModel.SampleDateTime_Local.Year
                                                                                 && c.DateTime_Local.Month == mwqmSampleAnalysisModel.SampleDateTime_Local.Month
                                                                                 && c.DateTime_Local.Day == mwqmSampleAnalysisModel.SampleDateTime_Local.Day
                                                                                 select c).FirstOrDefault();

                                    List<int> RainData = new List<int>()
                                        {
                                           Convert.ToInt32(mwqmRunAnalysisModel.RainDay0_mm),
                                           Convert.ToInt32(mwqmRunAnalysisModel.RainDay1_mm),
                                           Convert.ToInt32(mwqmRunAnalysisModel.RainDay2_mm),
                                           Convert.ToInt32(mwqmRunAnalysisModel.RainDay3_mm),
                                           Convert.ToInt32(mwqmRunAnalysisModel.RainDay4_mm),
                                           Convert.ToInt32(mwqmRunAnalysisModel.RainDay5_mm),
                                           Convert.ToInt32(mwqmRunAnalysisModel.RainDay6_mm),
                                           Convert.ToInt32(mwqmRunAnalysisModel.RainDay7_mm),
                                           Convert.ToInt32(mwqmRunAnalysisModel.RainDay8_mm),
                                           Convert.ToInt32(mwqmRunAnalysisModel.RainDay9_mm),
                                           Convert.ToInt32(mwqmRunAnalysisModel.RainDay10_mm),
                                        };

                                    int ShortRange = Math.Abs(mwqmAnalysisReportParameterModel.ShortRangeNumberOfDays);
                                    int MidRange = Math.Abs(mwqmAnalysisReportParameterModel.MidRangeNumberOfDays);
                                    int TotalRain = 0;
                                    bool AlreadyUsed = false;
                                    for (int i = 1; i < 11; i++)
                                    {
                                        TotalRain = TotalRain + RainData[i];
                                        if (i <= ShortRange)
                                        {
                                            if (i == 1)
                                            {
                                                if (mwqmAnalysisReportParameterModel.WetLimit24h <= TotalRain)
                                                {
                                                    int Col = 0;
                                                    for (int j = 0, count = mwqmSubsectorAnalysisModel.MWQMRunAnalysisModelList.Count(); j < count; j++)
                                                    {
                                                        if (mwqmSubsectorAnalysisModel.MWQMRunAnalysisModelList[j].MWQMRunID == mwqmRunAnalysisModel.MWQMRunID)
                                                        {
                                                            Col = j + 13;
                                                            break;
                                                        }
                                                    }
                                                    ws.Cells[3, Col].Select();
                                                    xlApp.Selection.Interior.Color = 16772300;

                                                    if (!AlreadyUsed)
                                                    {
                                                        mwqmSampleAnalysisForSiteModelToUseList.Add(mwqmSampleAnalysisModel);
                                                        AlreadyUsed = true;
                                                    }
                                                }
                                            }
                                            else if (i == 2)
                                            {
                                                if (mwqmAnalysisReportParameterModel.WetLimit48h <= TotalRain)
                                                {
                                                    int Col = 0;
                                                    for (int j = 0, count = mwqmSubsectorAnalysisModel.MWQMRunAnalysisModelList.Count(); j < count; j++)
                                                    {
                                                        if (mwqmSubsectorAnalysisModel.MWQMRunAnalysisModelList[j].MWQMRunID == mwqmRunAnalysisModel.MWQMRunID)
                                                        {
                                                            Col = j + 13;
                                                            break;
                                                        }
                                                    }
                                                    ws.Cells[4, Col].Select();
                                                    xlApp.Selection.Interior.Color = 16772300;

                                                    if (!AlreadyUsed)
                                                    {
                                                        mwqmSampleAnalysisForSiteModelToUseList.Add(mwqmSampleAnalysisModel);
                                                        AlreadyUsed = true;
                                                    }
                                                }
                                            }
                                            else if (i == 3)
                                            {
                                                if (mwqmAnalysisReportParameterModel.WetLimit72h <= TotalRain)
                                                {
                                                    int Col = 0;
                                                    for (int j = 0, count = mwqmSubsectorAnalysisModel.MWQMRunAnalysisModelList.Count(); j < count; j++)
                                                    {
                                                        if (mwqmSubsectorAnalysisModel.MWQMRunAnalysisModelList[j].MWQMRunID == mwqmRunAnalysisModel.MWQMRunID)
                                                        {
                                                            Col = j + 13;
                                                            break;
                                                        }
                                                    }
                                                    ws.Cells[5, Col].Select();
                                                    xlApp.Selection.Interior.Color = 16772300;

                                                    if (!AlreadyUsed)
                                                    {
                                                        mwqmSampleAnalysisForSiteModelToUseList.Add(mwqmSampleAnalysisModel);
                                                        AlreadyUsed = true;
                                                    }
                                                }
                                            }
                                            else
                                            {
                                                if (mwqmAnalysisReportParameterModel.WetLimit96h <= TotalRain)
                                                {
                                                    int Col = 0;
                                                    for (int j = 0, count = mwqmSubsectorAnalysisModel.MWQMRunAnalysisModelList.Count(); j < count; j++)
                                                    {
                                                        if (mwqmSubsectorAnalysisModel.MWQMRunAnalysisModelList[j].MWQMRunID == mwqmRunAnalysisModel.MWQMRunID)
                                                        {
                                                            Col = j + 13;
                                                            break;
                                                        }
                                                    }
                                                    ws.Cells[6, Col].Select();
                                                    xlApp.Selection.Interior.Color = 16772300;

                                                    if (!AlreadyUsed)
                                                    {
                                                        mwqmSampleAnalysisForSiteModelToUseList.Add(mwqmSampleAnalysisModel);
                                                        AlreadyUsed = true;
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                                else if (mwqmAnalysisReportParameterModel.AnalysisCalculationType == AnalysisCalculationTypeEnum.DryAllAll)
                                {
                                    MWQMRunAnalysisModel mwqmRunAnalysisModel = (from c in mwqmSubsectorAnalysisModel.MWQMRunAnalysisModelList
                                                                                 where c.DateTime_Local.Year == mwqmSampleAnalysisModel.SampleDateTime_Local.Year
                                                                                 && c.DateTime_Local.Month == mwqmSampleAnalysisModel.SampleDateTime_Local.Month
                                                                                 && c.DateTime_Local.Day == mwqmSampleAnalysisModel.SampleDateTime_Local.Day
                                                                                 select c).FirstOrDefault();

                                    List<int> RainData = new List<int>()
                                        {
                                           Convert.ToInt32(mwqmRunAnalysisModel.RainDay0_mm),
                                           Convert.ToInt32(mwqmRunAnalysisModel.RainDay1_mm),
                                           Convert.ToInt32(mwqmRunAnalysisModel.RainDay2_mm),
                                           Convert.ToInt32(mwqmRunAnalysisModel.RainDay3_mm),
                                           Convert.ToInt32(mwqmRunAnalysisModel.RainDay4_mm),
                                           Convert.ToInt32(mwqmRunAnalysisModel.RainDay5_mm),
                                           Convert.ToInt32(mwqmRunAnalysisModel.RainDay6_mm),
                                           Convert.ToInt32(mwqmRunAnalysisModel.RainDay7_mm),
                                           Convert.ToInt32(mwqmRunAnalysisModel.RainDay8_mm),
                                           Convert.ToInt32(mwqmRunAnalysisModel.RainDay9_mm),
                                           Convert.ToInt32(mwqmRunAnalysisModel.RainDay10_mm),
                                        };

                                    int ShortRange = Math.Abs(mwqmAnalysisReportParameterModel.ShortRangeNumberOfDays);
                                    int MidRange = Math.Abs(mwqmAnalysisReportParameterModel.MidRangeNumberOfDays);
                                    int TotalRain = 0;
                                    bool CanUsed = true;
                                    for (int i = 1; i < 11; i++)
                                    {
                                        TotalRain = TotalRain + RainData[i];
                                        if (i <= ShortRange)
                                        {
                                            if (i == 1)
                                            {
                                                if (mwqmAnalysisReportParameterModel.DryLimit24h < TotalRain)
                                                {
                                                    int Col = 0;
                                                    for (int j = 0, count = mwqmSubsectorAnalysisModel.MWQMRunAnalysisModelList.Count(); j < count; j++)
                                                    {
                                                        if (mwqmSubsectorAnalysisModel.MWQMRunAnalysisModelList[j].MWQMRunID == mwqmRunAnalysisModel.MWQMRunID)
                                                        {
                                                            Col = j + 13;
                                                            break;
                                                        }
                                                    }
                                                    ws.Cells[3, Col].Select();
                                                    xlApp.Selection.Interior.Color = 10079487;

                                                    CanUsed = false;
                                                }
                                            }
                                            else if (i == 2)
                                            {
                                                if (mwqmAnalysisReportParameterModel.DryLimit48h < TotalRain)
                                                {
                                                    int Col = 0;
                                                    for (int j = 0, count = mwqmSubsectorAnalysisModel.MWQMRunAnalysisModelList.Count(); j < count; j++)
                                                    {
                                                        if (mwqmSubsectorAnalysisModel.MWQMRunAnalysisModelList[j].MWQMRunID == mwqmRunAnalysisModel.MWQMRunID)
                                                        {
                                                            Col = j + 13;
                                                            break;
                                                        }
                                                    }
                                                    ws.Cells[4, Col].Select();
                                                    xlApp.Selection.Interior.Color = 10079487;

                                                    CanUsed = false;
                                                }
                                            }
                                            else if (i == 3)
                                            {
                                                if (mwqmAnalysisReportParameterModel.DryLimit72h < TotalRain)
                                                {
                                                    int Col = 0;
                                                    for (int j = 0, count = mwqmSubsectorAnalysisModel.MWQMRunAnalysisModelList.Count(); j < count; j++)
                                                    {
                                                        if (mwqmSubsectorAnalysisModel.MWQMRunAnalysisModelList[j].MWQMRunID == mwqmRunAnalysisModel.MWQMRunID)
                                                        {
                                                            Col = j + 13;
                                                            break;
                                                        }
                                                    }
                                                    ws.Cells[5, Col].Select();
                                                    xlApp.Selection.Interior.Color = 10079487;

                                                    CanUsed = false;
                                                }
                                            }
                                            else
                                            {
                                                if (mwqmAnalysisReportParameterModel.DryLimit96h < TotalRain)
                                                {
                                                    int Col = 0;
                                                    for (int j = 0, count = mwqmSubsectorAnalysisModel.MWQMRunAnalysisModelList.Count(); j < count; j++)
                                                    {
                                                        if (mwqmSubsectorAnalysisModel.MWQMRunAnalysisModelList[j].MWQMRunID == mwqmRunAnalysisModel.MWQMRunID)
                                                        {
                                                            Col = j + 13;
                                                            break;
                                                        }
                                                    }
                                                    ws.Cells[6, Col].Select();
                                                    xlApp.Selection.Interior.Color = 10079487;

                                                    CanUsed = false;
                                                }
                                            }
                                        }
                                    }
                                    if (CanUsed)
                                    {
                                        mwqmSampleAnalysisForSiteModelToUseList.Add(mwqmSampleAnalysisModel);
                                    }
                                }
                                else
                                {
                                    mwqmSampleAnalysisForSiteModelToUseList.Add(mwqmSampleAnalysisModel);
                                }
                            }
                        }
                    }
                }

                if (mwqmAnalysisReportParameterModel.AnalysisCalculationType == AnalysisCalculationTypeEnum.AllAllAll)
                {
                    if (mwqmSampleAnalysisForSiteModelToUseList.Count > 0 && mwqmAnalysisReportParameterModel.FullYear)
                    {
                        int FirstYear = mwqmSampleAnalysisForSiteModelToUseList[0].SampleDateTime_Local.Year;
                        int LastYear = mwqmSampleAnalysisForSiteModelToUseList[mwqmSampleAnalysisForSiteModelToUseList.Count - 1].SampleDateTime_Local.Year;

                        List<MWQMSampleAnalysisModel> mwqmSampleAnalysisMore = (from c in mwqmSampleAnalysisForSiteModelList
                                                                                where c.SampleDateTime_Local.Year == FirstYear
                                                                                && c.MWQMSiteTVItemID == mwqmSiteAnalysisModel.MWQMSiteTVItemID
                                                                                select c).Concat((from c in mwqmSampleAnalysisForSiteModelList
                                                                                                  where c.SampleDateTime_Local.Year == LastYear
                                                                                                  && c.MWQMSiteTVItemID == mwqmSiteAnalysisModel.MWQMSiteTVItemID
                                                                                                  select c)).ToList();

                        List<MWQMSampleAnalysisModel> mwqmSampleAnalysisMore2 = new List<MWQMSampleAnalysisModel>();
                        foreach (MWQMSampleAnalysisModel mwqmSampleAnalysisModel in mwqmSampleAnalysisMore)
                        {
                            if (!MWQMRunTVItemIDToOmitList.Contains(mwqmSampleAnalysisModel.MWQMRunTVItemID))
                            {
                                mwqmSampleAnalysisMore2.Add(mwqmSampleAnalysisModel);
                            }
                        }

                        mwqmSampleAnalysisForSiteModelToUseList = mwqmSampleAnalysisForSiteModelToUseList.Concat(mwqmSampleAnalysisMore2).Distinct().ToList();

                        mwqmSampleAnalysisForSiteModelToUseList = mwqmSampleAnalysisForSiteModelToUseList.OrderByDescending(c => c.SampleDateTime_Local).ToList();
                    }
                }

                int Coloring = 0;
                string Letter = "";
                if (mwqmSampleAnalysisForSiteModelToUseList.Count < 10)
                {
                    range = ws.Cells[RowCount, 3];
                    range.Value = "--";

                    range = ws.Cells[RowCount, 4];
                    range.Value = "--";

                    range = ws.Cells[RowCount, 5];
                    range.Value = "--";

                    range = ws.Cells[RowCount, 6];
                    range.Value = "--";

                    range = ws.Cells[RowCount, 7];
                    range.Value = "--";

                    range = ws.Cells[RowCount, 8];
                    range.Value = "--";

                    range = ws.Cells[RowCount, 9];
                    range.Value = "--";

                    range = ws.Cells[RowCount, 10];
                    range.Value = "--";

                    range = ws.Cells[RowCount, 11];
                    range.Value = "--";

                    Letter = mwqmSampleAnalysisForSiteModelToUseList.Count.ToString();
                    Coloring = 16764057;

                    if (mwqmSiteAnalysisModel.IsActive)
                    {
                        range = ws.Cells[RowCount, 12];
                        range.Value = "'" + Letter;
                        range.Select();
                        xlApp.Selection.Interior.Color = Coloring;
                    }
                }
                if (mwqmSampleAnalysisForSiteModelToUseList.Count >= 10)
                {
                    int MWQMSampleCount = mwqmSampleAnalysisForSiteModelToUseList.Count;
                    int? MaxYear = mwqmSampleAnalysisForSiteModelToUseList[0].SampleDateTime_Local.Year;
                    int? MinYear = mwqmSampleAnalysisForSiteModelToUseList[mwqmSampleAnalysisForSiteModelToUseList.Count - 1].SampleDateTime_Local.Year;
                    int? MinFC = (from c in mwqmSampleAnalysisForSiteModelToUseList select c.FecCol_MPN_100ml).Min();
                    int? MaxFC = (from c in mwqmSampleAnalysisForSiteModelToUseList select c.FecCol_MPN_100ml).Max();

                    List<double> SampleList = (from c in mwqmSampleAnalysisForSiteModelToUseList
                                               select (c.FecCol_MPN_100ml == 1 ? 1.9D : (double)c.FecCol_MPN_100ml)).ToList<double>();

                    P90 = tvItemService.GetP90(SampleList);
                    GeoMean = tvItemService.GeometricMean(SampleList);
                    Median = tvItemService.GetMedian(SampleList);
                    PercOver43 = ((((double)SampleList.Where(c => c > 43).Count()) / (double)SampleList.Count()) * 100.0D);
                    PercOver260 = ((((double)SampleList.Where(c => c > 260).Count()) / (double)SampleList.Count()) * 100.0D);
                    if ((GeoMean > 88) || (Median > 88) || (P90 > 260) || (PercOver260 > 10))
                    {
                        if ((GeoMean > 181.33) || (Median > 181.33) || (P90 > 460.0) || (PercOver260 > 18.33))
                        {
                            Coloring = 16746632;
                            Letter = "F";
                        }
                        else if ((GeoMean > 162.67) || (Median > 162.67) || (P90 > 420.0) || (PercOver260 > 16.67))
                        {
                            Coloring = 16751001;
                            Letter = "E";
                        }
                        else if ((GeoMean > 144.0) || (Median > 144.0) || (P90 > 380.0) || (PercOver260 > 15.0))
                        {
                            Coloring = 16755370;
                            Letter = "D";
                        }
                        else if ((GeoMean > 125.33) || (Median > 125.33) || (P90 > 340.0) || (PercOver260 > 13.33))
                        {
                            Coloring = 16759739;
                            Letter = "C";
                        }
                        else if ((GeoMean > 106.67) || (Median > 106.67) || (P90 > 300.0) || (PercOver260 > 11.67))
                        {
                            Coloring = 16764108;
                            Letter = "B";
                        }
                        else
                        {
                            Coloring = 16768477;
                            Letter = "A";
                        }
                    }
                    else if ((GeoMean > 14) || (Median > 14) || (P90 > 43) || (PercOver43 > 10))
                    {
                        if ((GeoMean > 75.67) || (Median > 75.67) || (P90 > 223.83) || (PercOver43 > 26.67))
                        {
                            Coloring = 170;
                            Letter = "F";
                        }
                        else if ((GeoMean > 63.33) || (Median > 63.33) || (P90 > 187.67) || (PercOver43 > 23.33))
                        {
                            Coloring = 204;
                            Letter = "E";
                        }
                        else if ((GeoMean > 51.0) || (Median > 51.0) || (P90 > 151.5) || (PercOver43 > 20.0))
                        {
                            Coloring = 1118718;
                            Letter = "D";
                        }
                        else if ((GeoMean > 38.67) || (Median > 38.67) || (P90 > 115.33) || (PercOver43 > 16.67))
                        {
                            Coloring = 4474111;
                            Letter = "C";
                        }
                        else if ((GeoMean > 26.33) || (Median > 26.33) || (P90 > 79.17) || (PercOver43 > 13.33))
                        {
                            Coloring = 10066431;
                            Letter = "B";
                        }
                        else
                        {
                            Coloring = 13421823;
                            Letter = "A";
                        }
                    }
                    else
                    {
                        if ((GeoMean > 11.67) || (Median > 11.67) || (P90 > 35.83) || (PercOver43 > 8.33))
                        {
                            Coloring = 13434828;
                            Letter = "F";
                        }
                        else if ((GeoMean > 9.33) || (Median > 9.33) || (P90 > 28.67) || (PercOver43 > 6.67))
                        {
                            Coloring = 10092441;
                            Letter = "E";
                        }
                        else if ((GeoMean > 7.0) || (Median > 7.0) || (P90 > 21.5) || (PercOver43 > 5.0))
                        {
                            Coloring = 4521796;
                            Letter = "D";
                        }
                        else if ((GeoMean > 4.67) || (Median > 4.67) || (P90 > 14.33) || (PercOver43 > 3.33))
                        {
                            Coloring = 1179409;
                            Letter = "C";
                        }
                        else if ((GeoMean > 2.33) || (Median > 2.33) || (P90 > 7.17) || (PercOver43 > 1.67))
                        {
                            Coloring = 47872;
                            Letter = "B";
                        }
                        else
                        {
                            Coloring = 39168;
                            Letter = "A";
                        }
                    }

                    range = ws.Cells[RowCount, 3];
                    range.Value = "'" + (mwqmSiteAnalysisModel.IsActive ? MWQMSampleCount.ToString() : "--");

                    range = ws.Cells[RowCount, 4];
                    range.Value = "'" + (mwqmSiteAnalysisModel.IsActive ? (MaxYear != null ? (MaxYear.ToString() + "-" + MinYear.ToString()) : "--") : "--");

                    range = ws.Cells[RowCount, 5];
                    range.Value = "'" + (mwqmSiteAnalysisModel.IsActive ? (MinFC != null ? (MinFC < 2 ? "< 2" : (MinFC.ToString())) : "--") : "--");

                    range = ws.Cells[RowCount, 6];
                    range.Value = "'" + (mwqmSiteAnalysisModel.IsActive ? (MaxFC != null ? (MaxFC < 2 ? "< 2" : (MaxFC.ToString())) : "--") : "--");

                    range = ws.Cells[RowCount, 7];
                    range.Value = "'" + (mwqmSiteAnalysisModel.IsActive ? (GeoMean != null ? ((double)GeoMean < 2.0D ? "< 2" : ((double)GeoMean).ToString("F0")) : "--") : "--");
                    if (GeoMean > 14)
                    {
                        if (mwqmSiteAnalysisModel.IsActive)
                        {
                            range.Interior.Color = 65535;
                        }
                    }

                    range = ws.Cells[RowCount, 8];
                    range.Value = "'" + (mwqmSiteAnalysisModel.IsActive ? (Median != null ? ((double)Median < 2.0D ? "< 2" : ((double)Median).ToString("F0")) : "--") : "--");
                    if (Median > 14)
                    {
                        if (mwqmSiteAnalysisModel.IsActive)
                        {
                            range.Interior.Color = 65535;
                        }
                    }

                    range = ws.Cells[RowCount, 9];
                    range.Value = "'" + (mwqmSiteAnalysisModel.IsActive ? (P90 != null ? ((double)P90 < 2.0D ? "< 2" : ((double)P90).ToString("F0")) : "--") : "--");
                    if (P90 > 43)
                    {
                        if (mwqmSiteAnalysisModel.IsActive)
                        {
                            range.Interior.Color = 65535;
                        }
                    }

                    range = ws.Cells[RowCount, 10];
                    range.Value = "'" + (mwqmSiteAnalysisModel.IsActive ? (PercOver43 != null ? ((double)PercOver43).ToString("F0") : "--") : "--");
                    if (PercOver43 > 10)
                    {
                        if (mwqmSiteAnalysisModel.IsActive)
                        {
                            range.Interior.Color = 65535;
                        }
                    }

                    range = ws.Cells[RowCount, 11];
                    range.Value = "'" + (mwqmSiteAnalysisModel.IsActive ? (PercOver260 != null ? ((double)PercOver260).ToString("F0") : "--") : "--");
                    if (PercOver260 > 10)
                    {
                        if (mwqmSiteAnalysisModel.IsActive)
                        {
                            range.Interior.Color = 65535;
                        }
                    }

                    if (mwqmSiteAnalysisModel.IsActive)
                    {
                        range = ws.Cells[RowCount, 12];
                        range.Value = "'" + Letter;
                        range.Select();
                        xlApp.Selection.Interior.Color = Coloring;
                    }
                }

                int AddedRows = 0;
                foreach (MWQMSampleAnalysisModel mwqmSampleAnalysisModel in mwqmSampleAnalysisForSiteModelList)
                {
                    AddedRows = 0;
                    int colNumber = (from c in runDateColNumberList
                                     where c.MWQMRunTVItemID == mwqmSampleAnalysisModel.MWQMRunTVItemID
                                     select c.ColNumber).FirstOrDefault();

                    if (mwqmSampleAnalysisForSiteModelToUseList.Count >= 10)
                    {
                        List<double> SampleList = (from c in mwqmSampleAnalysisForSiteModelToUseList
                                                   select (c.FecCol_MPN_100ml == 1 ? 1.9D : (double)c.FecCol_MPN_100ml)).ToList<double>();

                        P90 = tvItemService.GetP90(SampleList);
                        GeoMean = tvItemService.GeometricMean(SampleList);
                        Median = tvItemService.GetMedian(SampleList);
                        PercOver43 = ((((double)SampleList.Where(c => c > 43).Count()) / (double)SampleList.Count()) * 100.0D);
                        PercOver260 = ((((double)SampleList.Where(c => c > 260).Count()) / (double)SampleList.Count()) * 100.0D);

                    }
                    else
                    {
                        P90 = null;
                        GeoMean = null;
                        Median = null;
                        PercOver43 = null;
                        PercOver260 = null;
                    }

                    range = ws.Cells[RowCount, colNumber];

                    range.Value = (mwqmSampleAnalysisModel.FecCol_MPN_100ml < 2 ? "< 2" : mwqmSampleAnalysisModel.FecCol_MPN_100ml.ToString());
                    if (mwqmSampleAnalysisModel.FecCol_MPN_100ml > 500)
                    {
                        range.Interior.Color = 255;
                    }
                    else if (mwqmSampleAnalysisModel.FecCol_MPN_100ml > 43)
                    {
                        range.Interior.Color = 65535;
                    }
                    if (mwqmSiteAnalysisModel.IsActive == true)
                    {
                        if (mwqmSampleAnalysisForSiteModelToUseList.Contains(mwqmSampleAnalysisModel))
                        {
                            range.Select();
                            xlApp.Selection.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).Color = -11489280;
                            if (!RunUsedColNumberList.Contains(colNumber))
                            {
                                RunUsedColNumberList.Add(colNumber);
                            }
                        }
                        else
                        {
                            if (MWQMRunTVItemIDToOmitList.Contains(mwqmSampleAnalysisModel.MWQMRunTVItemID))
                            {
                                RunUsedColNumberList.Add(colNumber);
                            }
                        }
                    }
                    if (showDataTypeList.Contains(ExcelExportShowDataTypeEnum.Temperature))
                    {
                        RowCount += 1;
                        range = ws.Cells[RowCount, colNumber];
                        range.Value = (mwqmSampleAnalysisModel.WaterTemp_C != null ? ((double)mwqmSampleAnalysisModel.WaterTemp_C).ToString("F0") : "--");
                        if (!(rowAndTypeList.Where(c => c.RowNumber == (RowCount + AddedRows)).Any()))
                        {
                            rowAndTypeList.Add(new RowAndType() { RowNumber = RowCount + AddedRows, ExcelExportShowDataType = ExcelExportShowDataTypeEnum.Temperature });
                        }
                    }
                    if (showDataTypeList.Contains(ExcelExportShowDataTypeEnum.Salinity))
                    {
                        AddedRows += 1;
                        range = ws.Cells[RowCount + AddedRows, colNumber];
                        range.Value = (mwqmSampleAnalysisModel.Salinity_PPT != null ? ((double)mwqmSampleAnalysisModel.Salinity_PPT).ToString("F0") : "--");

                        if (mwqmSampleAnalysisModel.Salinity_PPT != null)
                        {
                            double? avgSal = (from c in mwqmSampleAnalysisForSiteModelList
                                              where c.Salinity_PPT != null
                                              select c.Salinity_PPT).Average();

                            if (Math.Abs(((double)mwqmSampleAnalysisModel.Salinity_PPT) - ((double)avgSal)) >= mwqmAnalysisReportParameterModel.SalinityHighlightDeviationFromAverage)
                            {
                                range.Interior.Color = 65535;
                            }
                        }
                        if (!(rowAndTypeList.Where(c => c.RowNumber == (RowCount + AddedRows)).Any()))
                        {
                            rowAndTypeList.Add(new RowAndType() { RowNumber = RowCount + AddedRows, ExcelExportShowDataType = ExcelExportShowDataTypeEnum.Salinity });
                        }

                    }
                    if (showDataTypeList.Contains(ExcelExportShowDataTypeEnum.P90))
                    {
                        AddedRows += 1;
                        range = ws.Cells[RowCount + AddedRows, colNumber];

                        range.Value = (P90 != null ? ((double)P90).ToString("F0") : "--");
                        if (P90 > 43)
                        {
                            range.Interior.Color = 65535;
                        }
                        if (!(rowAndTypeList.Where(c => c.RowNumber == (RowCount + AddedRows)).Any()))
                        {
                            rowAndTypeList.Add(new RowAndType() { RowNumber = RowCount + AddedRows, ExcelExportShowDataType = ExcelExportShowDataTypeEnum.P90 });
                        }
                    }
                    if (showDataTypeList.Contains(ExcelExportShowDataTypeEnum.GemetricMean))
                    {
                        AddedRows += 1;
                        range = ws.Cells[RowCount + AddedRows, colNumber];

                        range.Value = (GeoMean != null ? ((double)GeoMean).ToString("F0") : "--");
                        if (GeoMean > 14)
                        {
                            range.Interior.Color = 65535;
                        }
                        if (!(rowAndTypeList.Where(c => c.RowNumber == (RowCount + AddedRows)).Any()))
                        {
                            rowAndTypeList.Add(new RowAndType() { RowNumber = RowCount + AddedRows, ExcelExportShowDataType = ExcelExportShowDataTypeEnum.GemetricMean });
                        }
                    }
                    if (showDataTypeList.Contains(ExcelExportShowDataTypeEnum.Median))
                    {
                        AddedRows += 1;
                        range = ws.Cells[RowCount + AddedRows, colNumber];

                        range.Value = (Median != null ? ((double)Median).ToString("F0") : "--");
                        if (Median > 14)
                        {
                            range.Interior.Color = 65535;
                        }
                        if (!(rowAndTypeList.Where(c => c.RowNumber == (RowCount + AddedRows)).Any()))
                        {
                            rowAndTypeList.Add(new RowAndType() { RowNumber = RowCount + AddedRows, ExcelExportShowDataType = ExcelExportShowDataTypeEnum.Median });
                        }
                    }
                    if (showDataTypeList.Contains(ExcelExportShowDataTypeEnum.PercOfP90Over43))
                    {
                        AddedRows += 1;
                        range = ws.Cells[RowCount + AddedRows, colNumber];

                        range.Value = (PercOver43 != null ? ((double)PercOver43).ToString("F0") : "--");
                        if (PercOver43 > 20)
                        {
                            range.Interior.Color = 65535;
                        }
                        else if (PercOver43 > 10)
                        {
                            range.Interior.Color = 65535;
                        }
                        if (!(rowAndTypeList.Where(c => c.RowNumber == (RowCount + AddedRows)).Any()))
                        {
                            rowAndTypeList.Add(new RowAndType() { RowNumber = RowCount + AddedRows, ExcelExportShowDataType = ExcelExportShowDataTypeEnum.PercOfP90Over43 });
                        }
                    }
                    if (showDataTypeList.Contains(ExcelExportShowDataTypeEnum.PercOfP90Over260))
                    {
                        AddedRows += 1;
                        range = ws.Cells[RowCount + AddedRows, colNumber];

                        range.Value = (PercOver260 != null ? ((double)PercOver260).ToString("F0") : "--");
                        if (PercOver260 > 10)
                        {
                            range.Interior.Color = 65535;
                        }
                        if (!(rowAndTypeList.Where(c => c.RowNumber == (RowCount + AddedRows)).Any()))
                        {
                            rowAndTypeList.Add(new RowAndType() { RowNumber = RowCount + AddedRows, ExcelExportShowDataType = ExcelExportShowDataTypeEnum.PercOfP90Over260 });
                        }
                    }

                    if (mwqmSampleAnalysisForSiteModelToUseList.Contains(mwqmSampleAnalysisModel))
                    {
                        mwqmSampleAnalysisForSiteModelToUseList = mwqmSampleAnalysisForSiteModelToUseList.Skip(1).ToList();
                    }
                }

                RowCount += (1 + AddedRows);
            }

            foreach (RunDateColNumber runDateColNumber in runDateColNumberList)
            {
                if (MWQMRunTVItemIDToOmitList.Contains(runDateColNumber.MWQMRunTVItemID))
                {
                    if (RunUsedColNumberList.Contains(runDateColNumber.ColNumber))
                    {
                        range = ws.Cells[1, runDateColNumber.ColNumber];
                        range.Select();
                        xlApp.Selection.Interior.Color = 255;
                    }
                }
                else
                {
                    if (RunUsedColNumberList.Contains(runDateColNumber.ColNumber))
                    {
                        range = ws.Cells[1, runDateColNumber.ColNumber];
                        range.Select();
                        xlApp.Selection.Interior.Color = 5296274;
                    }
                }
            }

            xlApp.Range[ws.Cells[Math.Abs(mwqmAnalysisReportParameterModel.ShortRangeNumberOfDays) + 2, 13], ws.Cells[Math.Abs(mwqmAnalysisReportParameterModel.ShortRangeNumberOfDays) + 2, 13 + runDateColNumberList.Count() - 1]].Select();
            xlApp.Selection.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            xlApp.Selection.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).Color = -11489280;
            xlApp.Selection.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;

            xlApp.Range[ws.Cells[Math.Abs(mwqmAnalysisReportParameterModel.MidRangeNumberOfDays) + 2, 13], ws.Cells[Math.Abs(mwqmAnalysisReportParameterModel.MidRangeNumberOfDays) + 2, 13 + runDateColNumberList.Count() - 1]].Select();
            xlApp.Selection.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            xlApp.Selection.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).Color = -11489280;
            xlApp.Selection.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;

            foreach (RowAndType rowAndType in rowAndTypeList)
            {
                xlApp.Range["A" + rowAndType.RowNumber.ToString() + ":L" + rowAndType.RowNumber.ToString()].Select();
                xlApp.Selection.Merge();
                xlApp.Selection.HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlRight;
                switch (rowAndType.ExcelExportShowDataType)
                {
                    case ExcelExportShowDataTypeEnum.FecalColiform:
                    case ExcelExportShowDataTypeEnum.Temperature:
                    case ExcelExportShowDataTypeEnum.Salinity:
                    case ExcelExportShowDataTypeEnum.P90:
                    case ExcelExportShowDataTypeEnum.GemetricMean:
                    case ExcelExportShowDataTypeEnum.Median:
                    case ExcelExportShowDataTypeEnum.PercOfP90Over43:
                    case ExcelExportShowDataTypeEnum.PercOfP90Over260:
                        {
                            xlApp.Selection.Value = _BaseEnumService.GetEnumText_ExcelExportShowDataTypeEnum(rowAndType.ExcelExportShowDataType);
                        }
                        break;
                    default:
                        {
                            xlApp.Selection.Value = "Error";
                        }
                        break;
                }
            }

            xlApp.Range["M16:M16"].Select();
            xlApp.ActiveWindow.FreezePanes = true;

            ws.Range["A1"].Select();
        }
        #endregion Functions

    }
}
