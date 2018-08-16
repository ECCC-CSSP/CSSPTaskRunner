using CSSPEnums;
using CSSPModels;
using CSSPTaskRunnerHelper.Resources;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Principal;
using System.Text;
using System.Threading.Tasks;
using WQMS.RTNotifiers;

namespace CSSPTaskRunnerHelper
{
    public partial class TaskRunnerHelper : TaskRunnerHelperBase
    {
        #region Variables
        private int LabSheetLookupDelay = 4; // seconds
        private int TaskStatusOfRunnningLookupDelay = 5; // seconds
        private List<int> ShawnDonohueEmailTimeHourList = new List<int>() { 6, 12, 18 }; // hours to send email every day
        private int MPNLimitForEmail = 500;
        //private bool testing = false;
        #endregion Variables

        #region Properties
        private int SkipTimerCount { get; set; }
        private LanguageEnum LanguageRequest { get; set; }
        private IPrincipal User { get; set; }
        private List<TextLanguage> TextLanguageList { get; set; }
        #endregion Properties

        #region Constructors
        public TaskRunnerHelper() : base()
        {
            SkipTimerCount = 0;
            User = new GenericPrincipal(new GenericIdentity("Charles.LeBlanc2@Canada.ca", "Forms"), null);
            LanguageRequest = LanguageEnum.en;
            TextLanguageList = new List<TextLanguage>();
        }
        #endregion Constructors

        #region Functions public
        private void ExecuteTask(AppTask appTask)
        {
            switch (_BWObj.appTaskModel.AppTaskCommand)
            {
                case AppTaskCommandEnum.OpenDataCSVOfMWQMSites:
                    {
                        AppTaskService appTaskService = new AppTaskService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
                        AppTaskModel appTaskModel = appTaskService.GetAppTaskModelWithAppTaskIDDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID);
                        TxtService txtService = new TxtService(_TaskRunnerBaseService);
                        txtService.CreateCSVOfMWQMSites();
                        if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0)
                        {
                            appTaskService.PostDeleteAppTaskDB(appTaskModel.AppTaskID);
                        }
                        else
                        {
                            SendErrorTextToDB(_TaskRunnerBaseService._BWObj.TextLanguageList);
                        }
                    }
                    break;
                case AppTaskCommandEnum.OpenDataCSVNationalOfMWQMSites:
                    {
                        AppTaskService appTaskService = new AppTaskService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
                        AppTaskModel appTaskModel = appTaskService.GetAppTaskModelWithAppTaskIDDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID);
                        TxtService txtService = new TxtService(_TaskRunnerBaseService);
                        txtService.CreateCSVNationalOfMWQMSites();
                        if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0)
                        {
                            appTaskService.PostDeleteAppTaskDB(appTaskModel.AppTaskID);
                        }
                        else
                        {
                            SendErrorTextToDB(_TaskRunnerBaseService._BWObj.TextLanguageList);
                        }
                    }
                    break;
                case AppTaskCommandEnum.OpenDataKMZOfMWQMSites:
                    {
                        AppTaskService appTaskService = new AppTaskService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
                        AppTaskModel appTaskModel = appTaskService.GetAppTaskModelWithAppTaskIDDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID);
                        KmzService kmzService = new KmzService(_TaskRunnerBaseService);
                        kmzService.CreateKMZOfMWQMSites();
                        if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0)
                        {
                            appTaskService.PostDeleteAppTaskDB(appTaskModel.AppTaskID);
                        }
                        else
                        {
                            SendErrorTextToDB(_TaskRunnerBaseService._BWObj.TextLanguageList);
                        }
                    }
                    break;
                case AppTaskCommandEnum.OpenDataCSVOfMWQMSamples:
                    {
                        AppTaskService appTaskService = new AppTaskService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
                        AppTaskModel appTaskModel = appTaskService.GetAppTaskModelWithAppTaskIDDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID);
                        TxtService txtService = new TxtService(_TaskRunnerBaseService);
                        txtService.CreateCSVOfMWQMSamples();
                        if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0)
                        {
                            appTaskService.PostDeleteAppTaskDB(appTaskModel.AppTaskID);
                        }
                        else
                        {
                            SendErrorTextToDB(_TaskRunnerBaseService._BWObj.TextLanguageList);
                        }
                    }
                    break;
                case AppTaskCommandEnum.OpenDataCSVNationalOfMWQMSamples:
                    {
                        AppTaskService appTaskService = new AppTaskService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
                        AppTaskModel appTaskModel = appTaskService.GetAppTaskModelWithAppTaskIDDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID);
                        TxtService txtService = new TxtService(_TaskRunnerBaseService);
                        txtService.CreateCSVNationalOfMWQMSamples();
                        if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0)
                        {
                            appTaskService.PostDeleteAppTaskDB(appTaskModel.AppTaskID);
                        }
                        else
                        {
                            SendErrorTextToDB(_TaskRunnerBaseService._BWObj.TextLanguageList);
                        }
                    }
                    break;
                case AppTaskCommandEnum.OpenDataXlsxOfMWQMSitesAndSamples:
                    {
                        AppTaskService appTaskService = new AppTaskService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
                        AppTaskModel appTaskModel = appTaskService.GetAppTaskModelWithAppTaskIDDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID);
                        XlsxService xlsxService = new XlsxService(_TaskRunnerBaseService);
                        xlsxService.CreateXlsxOfMWQMSitesAndSamples();
                        if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0)
                        {
                            appTaskService.PostDeleteAppTaskDB(appTaskModel.AppTaskID);
                        }
                        else
                        {
                            SendErrorTextToDB(_TaskRunnerBaseService._BWObj.TextLanguageList);
                        }
                    }
                    break;
                case AppTaskCommandEnum.CreateWebTideDataWLAtFirstNode:
                    {
                        AppTaskService appTaskService = new AppTaskService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
                        AppTaskModel appTaskModel = appTaskService.GetAppTaskModelWithAppTaskIDDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID);
                        MikeScenarioFileService mikeScenarioFileService = new MikeScenarioFileService(_TaskRunnerBaseService);
                        mikeScenarioFileService.CreateWebTideDataWLAtFirstNode();
                        if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0)
                        {
                            appTaskService.PostDeleteAppTaskDB(appTaskModel.AppTaskID);
                        }
                        else
                        {
                            SendErrorTextToDB(_TaskRunnerBaseService._BWObj.TextLanguageList);
                        }
                    }
                    break;
                case AppTaskCommandEnum.ExportEmailDistributionLists:
                    {
                        AppTaskService appTaskService = new AppTaskService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
                        AppTaskModel appTaskModel = appTaskService.GetAppTaskModelWithAppTaskIDDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID);
                        XlsxService xlsxService = new XlsxService(_TaskRunnerBaseService);
                        xlsxService.ExportEmailDistributionLists();
                        if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0)
                        {
                            appTaskService.PostDeleteAppTaskDB(appTaskModel.AppTaskID);
                        }
                        else
                        {
                            SendErrorTextToDB(_TaskRunnerBaseService._BWObj.TextLanguageList);
                        }
                    }
                    break;
                case AppTaskCommandEnum.GenerateWebTide:
                    {
                        AppTaskService appTaskService = new AppTaskService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
                        AppTaskModel appTaskModel = appTaskService.GetAppTaskModelWithAppTaskIDDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID);
                        TidesAndCurrentsService tidesAndCurrentsService = new TidesAndCurrentsService(_TaskRunnerBaseService);
                        tidesAndCurrentsService.GenerateWebTideNodes();
                        if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0)
                        {
                            appTaskService.PostDeleteAppTaskDB(appTaskModel.AppTaskID);
                        }
                        else
                        {
                            SendErrorTextToDB(_TaskRunnerBaseService._BWObj.TextLanguageList);
                        }
                    }
                    break;
                case AppTaskCommandEnum.MikeScenarioAskToRun:
                    {
                        AppTaskService appTaskService = new AppTaskService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
                        AppTaskModel appTaskModel = appTaskService.GetAppTaskModelWithAppTaskIDDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID);
                        MikeScenarioFileService mikeScenarioFileService = new MikeScenarioFileService(_TaskRunnerBaseService);
                        mikeScenarioFileService.MikeScenarioAskToRun();
                        if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0)
                        {
                            //appTaskService.PostDeleteAppTaskDB(appTaskModel.AppTaskID);
                        }
                        else
                        {
                            SendErrorTextToDB(_TaskRunnerBaseService._BWObj.TextLanguageList);
                        }
                    }
                    break;
                case AppTaskCommandEnum.MikeScenarioWaitingToRun:
                    {
                        AppTaskService appTaskService = new AppTaskService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
                        AppTaskModel appTaskModel = appTaskService.GetAppTaskModelWithAppTaskIDDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID);
                        MikeScenarioFileService mikeScenarioFileService = new MikeScenarioFileService(_TaskRunnerBaseService);
                        //mikeScenarioFileService.MikeScenarioAskToRun();
                        if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0)
                        {
                            //appTaskService.PostDeleteAppTaskDB(appTaskModel.AppTaskID);
                        }
                        else
                        {
                            SendErrorTextToDB(_TaskRunnerBaseService._BWObj.TextLanguageList);
                        }
                    }
                    break;
                case AppTaskCommandEnum.MikeScenarioImport:
                    {
                        AppTaskService appTaskService = new AppTaskService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
                        AppTaskModel appTaskModel = appTaskService.GetAppTaskModelWithAppTaskIDDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID);
                        MikeScenarioFileService mikeScenarioFileService = new MikeScenarioFileService(_TaskRunnerBaseService);
                        mikeScenarioFileService.MikeScenarioImportDB();
                        if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0)
                        {
                            appTaskService.PostDeleteAppTaskDB(appTaskModel.AppTaskID);
                        }
                        else
                        {
                            SendErrorTextToDB(_TaskRunnerBaseService._BWObj.TextLanguageList);
                        }
                    }
                    break;
                case AppTaskCommandEnum.MikeScenarioOtherFileImport:
                    {
                        AppTaskService appTaskService = new AppTaskService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
                        AppTaskModel appTaskModel = appTaskService.GetAppTaskModelWithAppTaskIDDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID);
                        MikeScenarioFileService mikeScenarioFileService = new MikeScenarioFileService(_TaskRunnerBaseService);
                        mikeScenarioFileService.MikeScenarioOtherFileImportDB();
                        if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0)
                        {
                            appTaskService.PostDeleteAppTaskDB(appTaskModel.AppTaskID);
                        }
                        else
                        {
                            SendErrorTextToDB(_TaskRunnerBaseService._BWObj.TextLanguageList);
                        }
                    }
                    break;
                case AppTaskCommandEnum.MikeScenarioRunning:
                    {
                        // nothing
                    }
                    break;
                case AppTaskCommandEnum.SetupWebTide:
                    {
                        AppTaskService appTaskService = new AppTaskService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
                        AppTaskModel appTaskModel = appTaskService.GetAppTaskModelWithAppTaskIDDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID);
                        TidesAndCurrentsService tidesAndCurrentsService = new TidesAndCurrentsService(_TaskRunnerBaseService);
                        tidesAndCurrentsService.SetupWebTide();
                        if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0)
                        {
                            appTaskService.PostDeleteAppTaskDB(appTaskModel.AppTaskID);
                        }
                        else
                        {
                            SendErrorTextToDB(_TaskRunnerBaseService._BWObj.TextLanguageList);
                        }
                    }
                    break;
                case AppTaskCommandEnum.ExportToArcGIS:
                    {
                        AppTaskService appTaskService = new AppTaskService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
                        AppTaskModel appTaskModel = appTaskService.GetAppTaskModelWithAppTaskIDDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID);
                        XlsxService xlsxService = new XlsxService(_TaskRunnerBaseService);
                        xlsxService.ExportToArcGIS();
                        if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0)
                        {
                            appTaskService.PostDeleteAppTaskDB(appTaskModel.AppTaskID);
                        }
                        else
                        {
                            SendErrorTextToDB(_TaskRunnerBaseService._BWObj.TextLanguageList);
                        }
                    }
                    break;
                case AppTaskCommandEnum.GenerateClassificationForCSSPWebToolsVisualization:
                    {
                        AppTaskService appTaskService = new AppTaskService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
                        AppTaskModel appTaskModel = appTaskService.GetAppTaskModelWithAppTaskIDDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID);
                        ClassificationGenerateService classificationGenerateService = new ClassificationGenerateService(_TaskRunnerBaseService);
                        classificationGenerateService.GenerateClassificationForCSSPWebToolsVisualization();
                        if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0)
                        {
                            appTaskService.PostDeleteAppTaskDB(appTaskModel.AppTaskID);
                        }
                        else
                        {
                            SendErrorTextToDB(_TaskRunnerBaseService._BWObj.TextLanguageList);
                        }
                    }
                    break;
                case AppTaskCommandEnum.GenerateLinksBetweenMWQMSitesAndPolSourceSitesForCSSPWebToolsVisualization:
                    {
                        AppTaskService appTaskService = new AppTaskService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
                        AppTaskModel appTaskModel = appTaskService.GetAppTaskModelWithAppTaskIDDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID);
                        ClassificationGenerateService classificationGenerateService = new ClassificationGenerateService(_TaskRunnerBaseService);
                        classificationGenerateService.GenerateLinksBetweenMWQMSitesAndPolSourceSitesForCSSPWebToolsVisualization();
                        if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0)
                        {
                            appTaskService.PostDeleteAppTaskDB(appTaskModel.AppTaskID);
                        }
                        else
                        {
                            SendErrorTextToDB(_TaskRunnerBaseService._BWObj.TextLanguageList);
                        }
                    }
                    break;
                case AppTaskCommandEnum.GetAllPrecipitationForYear:
                    {
                        AppTaskService appTaskService = new AppTaskService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
                        AppTaskModel appTaskModel = appTaskService.GetAppTaskModelWithAppTaskIDDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID);
                        ClimateService climateService = new ClimateService(_TaskRunnerBaseService);
                        climateService.GetAllPrecipitationForYear();
                        if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0)
                        {
                            appTaskService.PostDeleteAppTaskDB(appTaskModel.AppTaskID);
                        }
                        else
                        {
                            SendErrorTextToDB(_TaskRunnerBaseService._BWObj.TextLanguageList);
                        }
                    }
                    break;
                case AppTaskCommandEnum.FillRunPrecipByClimateSitePriorityForYear:
                    {
                        AppTaskService appTaskService = new AppTaskService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
                        AppTaskModel appTaskModel = appTaskService.GetAppTaskModelWithAppTaskIDDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID);
                        ClimateService climateService = new ClimateService(_TaskRunnerBaseService);
                        climateService.FillRunPrecipByClimateSitePriorityForYear();
                        if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0)
                        {
                            appTaskService.PostDeleteAppTaskDB(appTaskModel.AppTaskID);
                        }
                        else
                        {
                            SendErrorTextToDB(_TaskRunnerBaseService._BWObj.TextLanguageList);
                        }
                    }
                    break;
                case AppTaskCommandEnum.FindMissingPrecipForProvince:
                    {
                        AppTaskService appTaskService = new AppTaskService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
                        AppTaskModel appTaskModel = appTaskService.GetAppTaskModelWithAppTaskIDDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID);
                        ClimateService climateService = new ClimateService(_TaskRunnerBaseService);
                        climateService.FindMissingPrecipForProvince();
                        if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0)
                        {
                            appTaskService.PostDeleteAppTaskDB(appTaskModel.AppTaskID);
                        }
                        else
                        {
                            SendErrorTextToDB(_TaskRunnerBaseService._BWObj.TextLanguageList);
                        }
                    }
                    break;
                case AppTaskCommandEnum.GetClimateSitesDataForRunsOfYear:
                    {
                        AppTaskService appTaskService = new AppTaskService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
                        AppTaskModel appTaskModel = appTaskService.GetAppTaskModelWithAppTaskIDDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID);
                        ClimateService climateService = new ClimateService(_TaskRunnerBaseService);
                        climateService.GetClimateSitesDataForRunsOfYear();
                        if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0)
                        {
                            appTaskService.PostDeleteAppTaskDB(appTaskModel.AppTaskID);
                        }
                        else
                        {
                            SendErrorTextToDB(_TaskRunnerBaseService._BWObj.TextLanguageList);
                        }
                    }
                    break;
                case AppTaskCommandEnum.UpdateClimateSiteInformation:
                    {
                        AppTaskService appTaskService = new AppTaskService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
                        AppTaskModel appTaskModel = appTaskService.GetAppTaskModelWithAppTaskIDDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID);
                        ClimateService climateService = new ClimateService(_TaskRunnerBaseService);
                        climateService.UpdateClimateSitesInformationForProvinceTVItemID();
                        if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0)
                        {
                            appTaskService.PostDeleteAppTaskDB(appTaskModel.AppTaskID);
                        }
                        else
                        {
                            SendErrorTextToDB(_TaskRunnerBaseService._BWObj.TextLanguageList);
                        }
                    }
                    break;
                case AppTaskCommandEnum.CreateFCForm:
                    {
                        AppTaskService appTaskService = new AppTaskService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
                        AppTaskModel appTaskModel = appTaskService.GetAppTaskModelWithAppTaskIDDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID);
                        DocxService docxService = new DocxService(_TaskRunnerBaseService);

                        string Parameters = _TaskRunnerBaseService._BWObj.appTaskModel.Parameters;
                        int LabSheetID = int.Parse(appTaskService.GetAppTaskParamStr(Parameters, "LabSheetID"));

                        LabSheetService labSheetService = new LabSheetService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
                        LabSheetModel labSheetModel = labSheetService.GetLabSheetModelWithLabSheetIDDB(LabSheetID);

                        SamplingPlanService samplingPlanService = new SamplingPlanService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
                        SamplingPlanModel samplingPlanModel = new SamplingPlanModel();
                        samplingPlanModel = samplingPlanService.GetSamplingPlanModelWithSamplingPlanIDDB(labSheetModel.SamplingPlanID);
                        //if (samplingPlanModel.IncludeLaboratoryQAQC)
                        //{
                        docxService.GenerateFCFormDOCX();
                        //}
                        if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0)
                        {
                            appTaskService.PostDeleteAppTaskDB(appTaskModel.AppTaskID);
                        }
                        else
                        {
                            SendErrorTextToDB(_TaskRunnerBaseService._BWObj.TextLanguageList);
                        }
                    }
                    break;
                case AppTaskCommandEnum.CreateSamplingPlanSamplingPlanTextFile:
                    {
                        AppTaskService appTaskService = new AppTaskService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
                        AppTaskModel appTaskModel = appTaskService.GetAppTaskModelWithAppTaskIDDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID);
                        TxtService txtService = new TxtService(_TaskRunnerBaseService);
                        txtService.CreateSamplingPlanSamplingPlan();
                        if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0)
                        {
                            appTaskService.PostDeleteAppTaskDB(appTaskModel.AppTaskID);
                        }
                        else
                        {
                            SendErrorTextToDB(_TaskRunnerBaseService._BWObj.TextLanguageList);
                        }
                    }
                    break;
                case AppTaskCommandEnum.ExportAnalysisToExcel:
                    {
                        AppTaskService appTaskService = new AppTaskService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
                        AppTaskModel appTaskModel = appTaskService.GetAppTaskModelWithAppTaskIDDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID);
                        XlsxService xlsxService = new XlsxService(_TaskRunnerBaseService);
                        xlsxService.CreateExcelFileForAnalysisReportParameter();
                        if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0)
                        {
                            appTaskService.PostDeleteAppTaskDB(appTaskModel.AppTaskID);
                        }
                        else
                        {
                            SendErrorTextToDB(_TaskRunnerBaseService._BWObj.TextLanguageList);
                        }
                    }
                    break;
                case AppTaskCommandEnum.CreateDocxPDF:
                    {
                        AppTaskService appTaskService = new AppTaskService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
                        AppTaskModel appTaskModel = appTaskService.GetAppTaskModelWithAppTaskIDDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID);
                        DocxService docxService = new DocxService(_TaskRunnerBaseService);
                        docxService.CreateDocxPDF();
                        if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0)
                        {
                            appTaskService.PostDeleteAppTaskDB(appTaskModel.AppTaskID);
                        }
                        else
                        {
                            SendErrorTextToDB(_TaskRunnerBaseService._BWObj.TextLanguageList);
                        }
                    }
                    break;
                case AppTaskCommandEnum.CreateXlsxPDF:
                    {
                        AppTaskService appTaskService = new AppTaskService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
                        AppTaskModel appTaskModel = appTaskService.GetAppTaskModelWithAppTaskIDDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID);
                        XlsxService xlsxService = new XlsxService(_TaskRunnerBaseService);
                        xlsxService.CreateXlsxPDF();
                        if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0)
                        {
                            appTaskService.PostDeleteAppTaskDB(appTaskModel.AppTaskID);
                        }
                        else
                        {
                            SendErrorTextToDB(_TaskRunnerBaseService._BWObj.TextLanguageList);
                        }
                    }
                    break;
                case AppTaskCommandEnum.CreateDocumentFromParameters:
                    {
                        AppTaskService appTaskService = new AppTaskService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
                        AppTaskModel appTaskModel = appTaskService.GetAppTaskModelWithAppTaskIDDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID);
                        ParametersService parametersService = new ParametersService(_TaskRunnerBaseService);
                        parametersService.Generate();
                        if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0)
                        {
                            appTaskService.PostDeleteAppTaskDB(appTaskModel.AppTaskID);
                        }
                        else
                        {
                            SendErrorTextToDB(_TaskRunnerBaseService._BWObj.TextLanguageList);
                        }
                    }
                    break;
                case AppTaskCommandEnum.CreateDocumentFromTemplate:
                    {
                        AppTaskService appTaskService = new AppTaskService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
                        AppTaskModel appTaskModel = appTaskService.GetAppTaskModelWithAppTaskIDDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID);
                        string[] ParamValueList = _TaskRunnerBaseService._BWObj.appTaskModel.Parameters.Split("|||".ToCharArray(), StringSplitOptions.RemoveEmptyEntries);

                        if (ParamValueList.Count() != 2)
                        {
                            string NotUsed = string.Format(TaskRunnerServiceRes.ParameterCount_NotEqual_, ParamValueList.Count().ToString(), "2");
                            _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("ParameterCount_NotEqual_", ParamValueList.Count().ToString(), "4");
                            if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0)
                            {
                                appTaskService.PostDeleteAppTaskDB(appTaskModel.AppTaskID);
                            }
                            else
                            {
                                SendErrorTextToDB(_TaskRunnerBaseService._BWObj.TextLanguageList);
                            }
                            return;
                        }

                        int TVItemID = (int.Parse(appTaskService.GetAppTaskParamStr(_TaskRunnerBaseService._BWObj.appTaskModel.Parameters, "TVItemID")));

                        if (TVItemID == 0)
                        {
                            string NotUsed = string.Format(TaskRunnerServiceRes.Parameter_NotFound, "TVItemID");
                            _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("Parameter_NotFound", "TVItemID");
                            if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0)
                            {
                                appTaskService.PostDeleteAppTaskDB(appTaskModel.AppTaskID);
                            }
                            else
                            {
                                SendErrorTextToDB(_TaskRunnerBaseService._BWObj.TextLanguageList);
                            }
                            return;
                        }

                        int DocTemplateID = (int.Parse(appTaskService.GetAppTaskParamStr(_TaskRunnerBaseService._BWObj.appTaskModel.Parameters, "DocTemplateID")));

                        if (DocTemplateID == 0)
                        {
                            string NotUsed = string.Format(TaskRunnerServiceRes.Parameter_NotFound, "DocTemplateID");
                            _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("Parameter_NotFound", "DocTemplateID");
                            if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0)
                            {
                                appTaskService.PostDeleteAppTaskDB(appTaskModel.AppTaskID);
                            }
                            else
                            {
                                SendErrorTextToDB(_TaskRunnerBaseService._BWObj.TextLanguageList);
                            }
                            return;
                        }

                        DocTemplateService docTemplateService = new DocTemplateService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
                        DocTemplateModel docTemplateModel = docTemplateService.GetDocTemplateModelWithDocTemplateIDDB(DocTemplateID);
                        if (!string.IsNullOrWhiteSpace(docTemplateModel.Error))
                        {
                            string NotUsed = string.Format(TaskRunnerServiceRes.CouldNotCreateDocumentFromTemplateError_, docTemplateModel.Error);
                            _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("CouldNotCreateDocumentFromTemplateError_", docTemplateModel.Error);
                            if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0)
                            {
                                appTaskService.PostDeleteAppTaskDB(appTaskModel.AppTaskID);
                            }
                            else
                            {
                                SendErrorTextToDB(_TaskRunnerBaseService._BWObj.TextLanguageList);
                            }
                            return;
                        }

                        if (string.IsNullOrWhiteSpace(docTemplateModel.FileName))
                        {
                            string NotUsed = string.Format(TaskRunnerServiceRes.FileNameForDocTemplateID_IsEmpty, DocTemplateID.ToString());
                            _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("FileNameForDocTemplateID_IsEmpty", DocTemplateID.ToString());
                            if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0)
                            {
                                appTaskService.PostDeleteAppTaskDB(appTaskModel.AppTaskID);
                            }
                            else
                            {
                                SendErrorTextToDB(_TaskRunnerBaseService._BWObj.TextLanguageList);
                            }
                            return;
                        }

                        TVFileService tvFileService = new TVFileService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
                        TVFileModel tvFileModel = tvFileService.GetTVFileModelWithTVFileTVItemIDDB(docTemplateModel.TVFileTVItemID);
                        if (!string.IsNullOrWhiteSpace(tvFileModel.Error))
                        {
                            string NotUsed = string.Format(TaskRunnerServiceRes.CouldNotCreateDocumentFromTemplateError_, tvFileModel.Error);
                            _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("CouldNotCreateDocumentFromTemplateError_", tvFileModel.Error);
                            if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0)
                            {
                                appTaskService.PostDeleteAppTaskDB(appTaskModel.AppTaskID);
                            }
                            else
                            {
                                SendErrorTextToDB(_TaskRunnerBaseService._BWObj.TextLanguageList);
                            }
                            return;
                        }

                        FileInfo fiTemplate = new FileInfo(tvFileModel.ServerFilePath + tvFileModel.ServerFileName);

                        if (!fiTemplate.Exists)
                        {
                            string NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFindFile_, fiTemplate.FullName);
                            _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("CouldNotFindFile_", fiTemplate.FullName);
                            if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0)
                            {
                                appTaskService.PostDeleteAppTaskDB(appTaskModel.AppTaskID);
                            }
                            else
                            {
                                SendErrorTextToDB(_TaskRunnerBaseService._BWObj.TextLanguageList);
                            }
                            return;
                        }

                        string NewFileName = fiTemplate.FullName.Replace("Template_", "");
                        string DateText = "_" + DateTime.Now.Year.ToString() + "_" + DateTime.Now.Month.ToString() + "_" + DateTime.Now.Day.ToString() + "_" + DateTime.Now.Hour.ToString() + "_" + DateTime.Now.Minute.ToString();

                        switch (fiTemplate.Extension.ToLower())
                        {
                            case ".csv":
                                NewFileName = NewFileName.Replace(".csv", DateText + ".csv");
                                break;
                            case ".docx":
                                NewFileName = NewFileName.Replace(".docx", DateText + ".docx");
                                break;
                            case ".xlsx":
                                NewFileName = NewFileName.Replace(".xlsx", DateText + ".xlsx");
                                break;
                            case ".kml":
                                NewFileName = NewFileName.Replace(".kml", DateText + ".kml");
                                break;
                            default:
                                break;
                        }

                        FileInfo fi = new FileInfo(NewFileName);

                        // need to create the file on the server
                        try
                        {
                            File.Copy(fiTemplate.FullName, fi.FullName);
                        }
                        catch (Exception ex)
                        {
                            string NotUsed = string.Format(TaskRunnerServiceRes.CouldNotCopyFile_To_Error_, fiTemplate.FullName, fi.FullName, ex.Message + " --- Inner: " + (ex.InnerException == null ? "" : ex.InnerException.Message));
                            _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("CouldNotCopyFile_To_Error_", fiTemplate.FullName, fi.FullName, ex.Message + " --- Inner: " + (ex.InnerException == null ? "" : ex.InnerException.Message));
                            if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0)
                            {
                                appTaskService.PostDeleteAppTaskDB(appTaskModel.AppTaskID);
                            }
                            else
                            {
                                SendErrorTextToDB(_TaskRunnerBaseService._BWObj.TextLanguageList);
                            }
                            return;
                        }

                        fi = new FileInfo(fi.FullName);
                        if (!fi.Exists)
                        {
                            string NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFindFile_, fi.FullName);
                            _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("CouldNotFindFile_", fi.FullName);
                            if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0)
                            {
                                appTaskService.PostDeleteAppTaskDB(appTaskModel.AppTaskID);
                            }
                            else
                            {
                                SendErrorTextToDB(_TaskRunnerBaseService._BWObj.TextLanguageList);
                            }
                            return;
                        }

                        switch (fiTemplate.Extension.ToLower())
                        {
                            case ".csv":
                                {
                                    ReportBaseService reportBaseService = new ReportBaseService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, new TreeView(), _TaskRunnerBaseService._User);
                                    string retStr = reportBaseService.GenerateReportFromTemplateCSV(fi, TVItemID, 99999999, _TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID);
                                    if (!string.IsNullOrWhiteSpace(retStr))
                                    {
                                        string NotUsed = string.Format(TaskRunnerServiceRes.CouldNotCreateDocumentFromTemplateError_, retStr);
                                        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("CouldNotCreateDocumentFromTemplateError_", retStr);
                                        SendErrorTextToDB(_TaskRunnerBaseService._BWObj.TextLanguageList);

                                        try
                                        {
                                            fi.Delete();
                                        }
                                        catch (Exception)
                                        {
                                            // nothing there is already a error message in the AppTask Table
                                        }
                                        return;
                                    }

                                    TVItemService tvItemService = new TVItemService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);

                                    TVItemModel tvItemModelExist = tvItemService.GetChildTVItemModelWithParentIDAndTVTextAndTVTypeDB(TVItemID, fi.Name, TVTypeEnum.File);
                                    if (string.IsNullOrWhiteSpace(tvItemModelExist.Error))
                                    {
                                        string NotUsed = string.Format(TaskRunnerServiceRes.FileAlreadyExist_, fi.Name);
                                        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("FileAlreadyExist_", fi.Name);
                                        if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0)
                                        {
                                            appTaskService.PostDeleteAppTaskDB(appTaskModel.AppTaskID);
                                        }
                                        else
                                        {
                                            SendErrorTextToDB(_TaskRunnerBaseService._BWObj.TextLanguageList);
                                        }
                                        return;
                                    }

                                    TVItemModel tvItemModelNew = tvItemService.PostAddChildTVItemDB(TVItemID, fi.Name, TVTypeEnum.File);
                                    if (!string.IsNullOrWhiteSpace(tvItemModelNew.Error))
                                    {
                                        string NotUsed = string.Format(TaskRunnerServiceRes.CouldNotAdd_Error_, TaskRunnerServiceRes.TVItem, tvItemModelNew.Error);
                                        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotAdd_Error_", TaskRunnerServiceRes.TVItem, tvItemModelNew.Error);
                                        if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0)
                                        {
                                            appTaskService.PostDeleteAppTaskDB(appTaskModel.AppTaskID);
                                        }
                                        else
                                        {
                                            SendErrorTextToDB(_TaskRunnerBaseService._BWObj.TextLanguageList);
                                        }
                                        return;
                                    }

                                    fi = new FileInfo(fi.FullName);

                                    TVFileModel tvFileModelNew = new TVFileModel()
                                    {
                                        TVFileTVItemID = tvItemModelNew.TVItemID,
                                        FilePurpose = FilePurposeEnum.TemplateGenerated,
                                        FileDescription = "Generated File",
                                        FileType = tvFileService.GetFileType(fi.Extension),
                                        FileSize_kb = Math.Max((int)fi.Length / 1024, 1),
                                        FileInfo = "Generated File",
                                        FileCreatedDate_UTC = DateTime.Now,
                                        FromWater = false,
                                        ClientFilePath = fi.Name,
                                        ServerFileName = fi.Name,
                                        ServerFilePath = fi.DirectoryName + "\\",
                                        Language = tvFileModel.Language,
                                        Year = DateTime.Now.Year,
                                    };

                                    TVFile tvFileExist = tvFileService.GetTVFileExistDB(tvFileModelNew);
                                    if (tvFileExist != null)
                                    {
                                        string NotUsed = string.Format(TaskRunnerServiceRes._AlreadyExist, TaskRunnerServiceRes.TVFile);
                                        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotAdd_Error_", TaskRunnerServiceRes.TVItem, tvFileModel.Error);
                                        if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0)
                                        {
                                            appTaskService.PostDeleteAppTaskDB(appTaskModel.AppTaskID);
                                        }
                                        else
                                        {
                                            SendErrorTextToDB(_TaskRunnerBaseService._BWObj.TextLanguageList);
                                        }
                                        return;
                                    }

                                    TVFileModel tvFileModelRet = tvFileService.PostAddTVFileDB(tvFileModelNew);
                                    if (!string.IsNullOrWhiteSpace(tvFileModelRet.Error))
                                    {
                                        string NotUsed = string.Format(TaskRunnerServiceRes.CouldNotAdd_Error_, TaskRunnerServiceRes.TVFile, tvFileModel.Error);
                                        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotAdd_Error_", TaskRunnerServiceRes.TVFile, tvFileModel.Error);
                                        if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0)
                                        {
                                            appTaskService.PostDeleteAppTaskDB(appTaskModel.AppTaskID);
                                        }
                                        else
                                        {
                                            SendErrorTextToDB(_TaskRunnerBaseService._BWObj.TextLanguageList);
                                        }
                                        return;
                                    }

                                    appTaskService.PostDeleteAppTaskDB(_BWObj.appTaskModel.AppTaskID);

                                }
                                break;
                            case ".docx":
                                {
                                    bool KeepAppTask = false;
                                    ReportBaseService reportBaseService = new ReportBaseService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, new TreeView(), _TaskRunnerBaseService._User);
                                    string retStr = reportBaseService.GenerateReportFromTemplateWord(fi, TVItemID, 99999999, _TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID);
                                    if (!string.IsNullOrWhiteSpace(retStr))
                                    {
                                        string NotUsed = string.Format(TaskRunnerServiceRes.ErrorsWhileCreatingDocumentDeleteTaskAndCheckCreatedDocument, fi.Name);
                                        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageList("ErrorsWhileCreatingDocumentDeleteTaskAndCheckCreatedDocument");
                                        if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0)
                                        {
                                            appTaskService.PostDeleteAppTaskDB(appTaskModel.AppTaskID);
                                        }
                                        else
                                        {
                                            SendErrorTextToDB(_TaskRunnerBaseService._BWObj.TextLanguageList);
                                        }
                                        KeepAppTask = true;
                                    }

                                    TVItemService tvItemService = new TVItemService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);

                                    TVItemModel tvItemModelExist = tvItemService.GetChildTVItemModelWithParentIDAndTVTextAndTVTypeDB(TVItemID, fi.Name, TVTypeEnum.File);
                                    if (string.IsNullOrWhiteSpace(tvItemModelExist.Error))
                                    {
                                        string NotUsed = string.Format(TaskRunnerServiceRes.FileAlreadyExist_, fi.Name);
                                        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("FileAlreadyExist_", fi.Name);
                                        if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0)
                                        {
                                            appTaskService.PostDeleteAppTaskDB(appTaskModel.AppTaskID);
                                        }
                                        else
                                        {
                                            SendErrorTextToDB(_TaskRunnerBaseService._BWObj.TextLanguageList);
                                        }
                                        return;
                                    }

                                    TVItemModel tvItemModelNew = tvItemService.PostAddChildTVItemDB(TVItemID, fi.Name, TVTypeEnum.File);
                                    if (!string.IsNullOrWhiteSpace(tvItemModelNew.Error))
                                    {
                                        string NotUsed = string.Format(TaskRunnerServiceRes.CouldNotAdd_Error_, TaskRunnerServiceRes.TVItem, tvItemModelNew.Error);
                                        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotAdd_Error_", TaskRunnerServiceRes.TVItem, tvItemModelNew.Error);
                                        if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0)
                                        {
                                            appTaskService.PostDeleteAppTaskDB(appTaskModel.AppTaskID);
                                        }
                                        else
                                        {
                                            SendErrorTextToDB(_TaskRunnerBaseService._BWObj.TextLanguageList);
                                        }
                                        return;
                                    }

                                    fi = new FileInfo(fi.FullName);

                                    TVFileModel tvFileModelNew = new TVFileModel()
                                    {
                                        TVFileTVItemID = tvItemModelNew.TVItemID,
                                        FilePurpose = FilePurposeEnum.TemplateGenerated,
                                        FileDescription = "Generated File",
                                        FileType = tvFileService.GetFileType(fi.Extension),
                                        FileSize_kb = Math.Max((int)fi.Length / 1024, 1),
                                        FileInfo = "Generated File",
                                        FileCreatedDate_UTC = DateTime.Now,
                                        FromWater = false,
                                        ClientFilePath = fi.Name,
                                        ServerFileName = fi.Name,
                                        ServerFilePath = fi.DirectoryName + "\\",
                                        Language = tvFileModel.Language,
                                        Year = DateTime.Now.Year,
                                    };

                                    TVFile tvFileExist = tvFileService.GetTVFileExistDB(tvFileModelNew);
                                    if (tvFileExist != null)
                                    {
                                        string NotUsed = string.Format(TaskRunnerServiceRes._AlreadyExist, TaskRunnerServiceRes.TVFile);
                                        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotAdd_Error_", TaskRunnerServiceRes.TVItem, tvFileModel.Error);
                                        if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0)
                                        {
                                            appTaskService.PostDeleteAppTaskDB(appTaskModel.AppTaskID);
                                        }
                                        else
                                        {
                                            SendErrorTextToDB(_TaskRunnerBaseService._BWObj.TextLanguageList);
                                        }
                                        return;
                                    }

                                    TVFileModel tvFileModelRet = tvFileService.PostAddTVFileDB(tvFileModelNew);
                                    if (!string.IsNullOrWhiteSpace(tvFileModelRet.Error))
                                    {
                                        string NotUsed = string.Format(TaskRunnerServiceRes.CouldNotAdd_Error_, TaskRunnerServiceRes.TVFile, tvFileModel.Error);
                                        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotAdd_Error_", TaskRunnerServiceRes.TVFile, tvFileModel.Error);
                                        if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0)
                                        {
                                            appTaskService.PostDeleteAppTaskDB(appTaskModel.AppTaskID);
                                        }
                                        else
                                        {
                                            SendErrorTextToDB(_TaskRunnerBaseService._BWObj.TextLanguageList);
                                        }
                                        return;
                                    }

                                    if (!KeepAppTask)
                                        appTaskService.PostDeleteAppTaskDB(_BWObj.appTaskModel.AppTaskID);

                                }
                                break;
                            case ".xlsx":
                                {
                                    //XlsxService xlsxService = new XlsxService(_TaskRunnerBaseService);
                                    //xlsxService.CreateDocumentFromTemplateXlsx(docTemplateModel);
                                    //UpdateAppTaskWithDeleteOrError();
                                    string NotUsed = string.Format(TaskRunnerServiceRes._NotImplemented, _BWObj.appTaskModel.AppTaskCommand.ToString());
                                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("_NotImplemented", _BWObj.appTaskModel.AppTaskCommand.ToString());
                                    if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0)
                                    {
                                        appTaskService.PostDeleteAppTaskDB(appTaskModel.AppTaskID);
                                    }
                                    else
                                    {
                                        SendErrorTextToDB(_TaskRunnerBaseService._BWObj.TextLanguageList);
                                    }
                                }
                                break;
                            case ".kml":
                                {
                                    bool KeepAppTask = false;
                                    ReportBaseService reportBaseService = new ReportBaseService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, new TreeView(), _TaskRunnerBaseService._User);
                                    string retStr = reportBaseService.GenerateReportFromTemplateKML(fi, TVItemID, 99999999, _TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID);
                                    if (!string.IsNullOrWhiteSpace(retStr))
                                    {
                                        string NotUsed = string.Format(TaskRunnerServiceRes.CouldNotCreateDocumentFromTemplateError_, retStr);
                                        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("CouldNotCreateDocumentFromTemplateError_", retStr);
                                        SendErrorTextToDB(_TaskRunnerBaseService._BWObj.TextLanguageList);

                                        try
                                        {
                                            fi.Delete();
                                        }
                                        catch (Exception)
                                        {
                                            // nothing there is already a error message in the AppTask Table
                                        }
                                        return;
                                    }

                                    TVItemService tvItemService = new TVItemService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);

                                    TVItemModel tvItemModelExist = tvItemService.GetChildTVItemModelWithParentIDAndTVTextAndTVTypeDB(TVItemID, fi.Name, TVTypeEnum.File);
                                    if (string.IsNullOrWhiteSpace(tvItemModelExist.Error))
                                    {
                                        string NotUsed = string.Format(TaskRunnerServiceRes.FileAlreadyExist_, fi.Name);
                                        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("FileAlreadyExist_", fi.Name);
                                        if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0)
                                        {
                                            appTaskService.PostDeleteAppTaskDB(appTaskModel.AppTaskID);
                                        }
                                        else
                                        {
                                            SendErrorTextToDB(_TaskRunnerBaseService._BWObj.TextLanguageList);
                                        }
                                        return;
                                    }

                                    TVItemModel tvItemModelNew = tvItemService.PostAddChildTVItemDB(TVItemID, fi.Name, TVTypeEnum.File);
                                    if (!string.IsNullOrWhiteSpace(tvItemModelNew.Error))
                                    {
                                        string NotUsed = string.Format(TaskRunnerServiceRes.CouldNotAdd_Error_, TaskRunnerServiceRes.TVItem, tvItemModelNew.Error);
                                        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotAdd_Error_", TaskRunnerServiceRes.TVItem, tvItemModelNew.Error);
                                        if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0)
                                        {
                                            appTaskService.PostDeleteAppTaskDB(appTaskModel.AppTaskID);
                                        }
                                        else
                                        {
                                            SendErrorTextToDB(_TaskRunnerBaseService._BWObj.TextLanguageList);
                                        }
                                        return;
                                    }

                                    fi = new FileInfo(fi.FullName);

                                    TVFileModel tvFileModelNew = new TVFileModel()
                                    {
                                        TVFileTVItemID = tvItemModelNew.TVItemID,
                                        FilePurpose = FilePurposeEnum.TemplateGenerated,
                                        FileDescription = "Generated File",
                                        FileType = tvFileService.GetFileType(fi.Extension),
                                        FileSize_kb = Math.Max((int)fi.Length / 1024, 1),
                                        FileInfo = "Generated File",
                                        FileCreatedDate_UTC = DateTime.Now,
                                        FromWater = false,
                                        ClientFilePath = fi.Name,
                                        ServerFileName = fi.Name,
                                        ServerFilePath = fi.DirectoryName + "\\",
                                        Language = tvFileModel.Language,
                                        Year = DateTime.Now.Year,
                                    };

                                    TVFile tvFileExist = tvFileService.GetTVFileExistDB(tvFileModelNew);
                                    if (tvFileExist != null)
                                    {
                                        string NotUsed = string.Format(TaskRunnerServiceRes._AlreadyExist, TaskRunnerServiceRes.TVFile);
                                        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotAdd_Error_", TaskRunnerServiceRes.TVItem, tvFileModel.Error);
                                        if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0)
                                        {
                                            appTaskService.PostDeleteAppTaskDB(appTaskModel.AppTaskID);
                                        }
                                        else
                                        {
                                            SendErrorTextToDB(_TaskRunnerBaseService._BWObj.TextLanguageList);
                                        }
                                        return;
                                    }

                                    TVFileModel tvFileModelRet = tvFileService.PostAddTVFileDB(tvFileModelNew);
                                    if (!string.IsNullOrWhiteSpace(tvFileModelRet.Error))
                                    {
                                        string NotUsed = string.Format(TaskRunnerServiceRes.CouldNotAdd_Error_, TaskRunnerServiceRes.TVFile, tvFileModel.Error);
                                        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotAdd_Error_", TaskRunnerServiceRes.TVFile, tvFileModel.Error);
                                        if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0)
                                        {
                                            appTaskService.PostDeleteAppTaskDB(appTaskModel.AppTaskID);
                                        }
                                        else
                                        {
                                            SendErrorTextToDB(_TaskRunnerBaseService._BWObj.TextLanguageList);
                                        }
                                        return;
                                    }

                                    if (!KeepAppTask)
                                        appTaskService.PostDeleteAppTaskDB(_BWObj.appTaskModel.AppTaskID);

                                }
                                break;
                            default:
                                break;
                        }

                        // need to add the file in TVFile and TVItemID

                    }
                    break;
                default:
                    {
                        AppTaskService appTaskService = new AppTaskService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
                        AppTaskModel appTaskModel = appTaskService.GetAppTaskModelWithAppTaskIDDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID);
                        string NotUsed = string.Format(TaskRunnerServiceRes._NotImplemented, _BWObj.appTaskModel.AppTaskCommand.ToString());
                        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("_NotImplemented", _BWObj.appTaskModel.AppTaskCommand.ToString());
                        if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0)
                        {
                            appTaskService.PostDeleteAppTaskDB(appTaskModel.AppTaskID);
                        }
                        else
                        {
                            SendErrorTextToDB(_TaskRunnerBaseService._BWObj.TextLanguageList);
                        }
                    }
                    break;
            }
        }
        public void GetNextTask()
        {
            if (SkipTimerCount == 86401)
            {
                SkipTimerCount = 0;
            }
            SkipTimerCount += 1;

            MessageEvent(new MessageEventArgs(DateTime.Now.ToString("F"), MessageTypeEnum.Status));

            using (CSSPWebToolsDBContext db = new CSSPWebToolsDBContext())
            {
                AppTask appTask = (from c in db.AppTasks
                                   where c.AppTaskStatus == AppTaskStatusEnum.Created
                                   select c).FirstOrDefault();

                if (appTask != null)
                {
                    appTask.AppTaskStatus = AppTaskStatusEnum.Running;

                    try
                    {
                        db.SaveChanges();
                    }
                    catch (Exception ex)
                    {
                        string InnerException = (ex.InnerException != null ? $" InnerException: { ex.InnerException.Message }" : "");
                        MessageEvent(new MessageEventArgs($"{ ex.Message }{ InnerException }", MessageTypeEnum.Status));
                        return;
                    }

                    if (Enum.GetNames(typeof(AppTaskCommandEnum)).Contains(appTask.AppTaskCommand.ToString()))
                    {
                        ExecuteTask(appTask);
                    }
                    else
                    {
                        MessageEvent(new MessageEventArgs($"Could not find command [{ appTask.AppTaskCommand.ToString() }]", MessageTypeEnum.Error));
                    }
                }
            }

            // Checking for new lab sheets
            if (SkipTimerCount % LabSheetLookupDelay == 0)
            {
                GetNextLabSheet();
            }

            #region commented out section
            // ****************************************************************************************
            // The next commented out section should be checked/updated using the CSSPTaskRunnerMIKE
            // ****************************************************************************************
            //// Checking others
            //if (SkipTimerCount % TaskStatusOfRunnningLookupDelay == 0) // if timer interval is 1 second then this will run every 5 seconds
            //{
            //    _TaskRunnerBaseService.UpdateStatusOfRunningScenarios();
            //}
            #endregion commented out section

            // Checking others
            DateTime CurrentDateTime = DateTime.Now;
            if (ShawnDonohueEmailTimeHourList.Contains(CurrentDateTime.Hour) && CurrentDateTime.Minute == 0 && CurrentDateTime.Second == 0)
            {
                try
                {
                    RTStationMonitor rtStationMonitor = new RTStationMonitor();

                    string ShawnDonohueErrorMessage = rtStationMonitor.CheckAutmatedStationsAndNotify();
                    if (!string.IsNullOrWhiteSpace(ShawnDonohueErrorMessage))
                    {
                        MessageEvent(new MessageEventArgs($"Shawn Donohue issue: { ShawnDonohueErrorMessage }", MessageTypeEnum.Error));
                    }
                }
                catch (Exception ex)
                {
                    string InnerException = (ex.InnerException == null ? "" : $" InnerException: { ex.InnerException.Message }");
                    MessageEvent(new MessageEventArgs($"Shawn Donohue issue: { ex.Message }{ InnerException }", MessageTypeEnum.Error));

                    ShawnDonohueEmail(ex);
                }
            }


            #region commented out section

            // ****************************************************************************************
            // The next commented out section was removed because it is not used and might potentially be removed for good
            // It's use is to inform biologist by sending them an email at 14:00 on Friday afternoon
            // if some lab sheet with their responsibility need to be approved
            // ****************************************************************************************
            //if (CurrentDateTime.DayOfWeek == DayOfWeek.Friday && CurrentDateTime.Hour == 14 && CurrentDateTime.Minute == 0 && CurrentDateTime.Second == 0)
            //{
            //    SamplingPlanService samplingPlanService = new SamplingPlanService(LanguageEnum.en, _User);

            //    List<SamplingPlanModel> samplingPlanModelList = samplingPlanService.GetAllActiveSamplingPlanModelListDB();

            //    List<int> ProvinceTVItemIDList = samplingPlanModelList.Select(c => c.ProvinceTVItemID).Distinct().ToList();

            //    foreach (int provinceTVItemID in ProvinceTVItemIDList)
            //    {
            //        TVItemService tvItemService = new TVItemService(LanguageEnum.en, _User);
            //        TVItemModel tvItemModelProv = tvItemService.GetTVItemModelWithTVItemIDDB(provinceTVItemID);

            //        List<string> ToEmailList = new List<string>();
            //        int TotalLabSheetTransferred = 0;
            //        foreach (SamplingPlanModel samplingPlanModel in samplingPlanModelList.Where(c => c.ProvinceTVItemID == provinceTVItemID && c.IsActive == true))
            //        {
            //            LabSheetService labSheetService = new LabSheetService(LanguageEnum.en, _User);

            //            TotalLabSheetTransferred += labSheetService.GetLabSheetCountWithSamplingPlanIDAndLabSheetStatusDB(samplingPlanModel.SamplingPlanID, LabSheetStatusEnum.Transferred);

            //            SamplingPlanEmailService samplingPlanEmailService = new SamplingPlanEmailService(LanguageEnum.en, _User);

            //            List<SamplingPlanEmailModel> SamplingPlanEmailModelList = samplingPlanEmailService.GetSamplingPlanEmailModelListWithSamplingPlanIDDB(samplingPlanModel.SamplingPlanID);

            //            foreach (SamplingPlanEmailModel samplingPlanEmailModel in SamplingPlanEmailModelList)
            //            {
            //                if (!samplingPlanEmailModel.IsContractor && samplingPlanEmailModel.FridayReminderAt14h)
            //                {
            //                    ToEmailList.Add(samplingPlanEmailModel.Email);
            //                }
            //            }
            //        }

            //        MailMessage mail = new MailMessage();

            //        foreach (string email in ToEmailList)
            //        {
            //            mail.To.Add(email.ToLower());
            //        }

            //        if (mail.To.Count == 0)
            //        {
            //            continue;
            //        }

            //        mail.From = new MailAddress("ec.pccsm-cssp.ec@canada.ca");
            //        mail.IsBodyHtml = true;

            //        SmtpClient myClient = new System.Net.Mail.SmtpClient();

            //        myClient.Host = "smtp.email-courriel.canada.ca";
            //        myClient.Port = 587;
            //        myClient.Credentials = new System.Net.NetworkCredential("ec.pccsm-cssp.ec@canada.ca", "H^9h6g@Gy$N57k=Dr@J7=F2y6p6b!T");
            //        myClient.EnableSsl = true;

            //        mail.Priority = MailPriority.High;

            //        string subject = $@"{TotalLabSheetTransferred} lab sheets waiting to be approved / {TotalLabSheetTransferred} de feuilles de laboratoire en attente approuvées";

            //        StringBuilder msg = new StringBuilder();

            //        // ---------------------- English part --------------

            //        string href = "http://wmon01dtchlebl2/csspwebtools/en-CA/#!View/" + tvItemModelProv.TVText + "|||" + tvItemModelProv.TVItemID.ToString() + "|||010003030200000000000000000000";

            //        msg.AppendLine(@"<p>(français suit)</p>");
            //        msg.AppendLine($@"<h2>{TotalLabSheetTransferred} lab sheets waiting to be approved</h2>");
            //        msg.AppendLine(@"<a href=""" + href + @""">Open CSSPWebTools</a>");
            //        msg.AppendLine(@"<br>");
            //        msg.AppendLine(@"<p>Auto email from CSSPWebTools</p>");
            //        msg.AppendLine(@"<br>");
            //        msg.AppendLine(@"<hr />");

            //        // ---------------------- French part --------------

            //        msg.AppendLine(@"<hr />");
            //        msg.AppendLine(@"<br>");
            //        msg.AppendLine($@"<h2>{TotalLabSheetTransferred} feuilles de laboratoire en attente d'approbation</h2>");
            //        msg.AppendLine(@"<a href=""" + href.Replace("en-CA", "fr-CA") + @""">Ouvrir CSSPWebTools</a>");
            //        msg.AppendLine(@"<br>");
            //        msg.AppendLine(@"<p>Courriel automatique provenant de CSSPWebTools</p>");

            //        mail.Subject = subject;
            //        mail.Body = msg.ToString();
            //        myClient.Send(mail);
            //    }
            //}

            #endregion commented out section

        }
        public void ShowTasksInformation()
        {
            MessageEvent(new MessageEventArgs("", MessageTypeEnum.Clear));
            MessageEvent(new MessageEventArgs("Should show all tasks information", MessageTypeEnum.Permanent));
        }
        #endregion Functions public

        #region Functions private
        #endregion Functions private
    }
}
