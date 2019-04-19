using System;
using System.Collections.Generic;
using Transfer.Utility;
using Transfer.ViewModels;
using static Transfer.Enum.Ref;

namespace Transfer.Models.Interface
{
    public interface ID6Repository
    {
        Tuple<bool, List<IFRS9_User>> getIFRS9_Assessment_Presented_Config(string groupProductCode, string tableId, string userAccount);
        Tuple<bool, List<IFRS9_User>> getIFRS9_Assessment_Auditor_Config(string groupProductCode, string tableId, string userAccount);

        Tuple<bool, List<D60ViewModel>> getD60All();

        Tuple<bool, List<D60ViewModel>> getD60(D60ViewModel dataModel);

        MSGReturnModel saveD60(string actionType, D60ViewModel dataModel);

        MSGReturnModel deleteD60(string parmID);

        MSGReturnModel sendD60ToAudit(string parmID, string auditor);

        MSGReturnModel D60Audit(string parmID, string status);

        Tuple<bool, List<D61ViewModel>> getD61All();

        Tuple<bool, List<D61ViewModel>> getD61(D61ViewModel dataModel);

        MSGReturnModel saveD61(string actionType, D61ViewModel dataModel);

        MSGReturnModel deleteD61(string checkItemCode, string Id);

        MSGReturnModel sendD61ToAudit(List<D61ViewModel> dataModelList);

        MSGReturnModel D61Audit(List<D61ViewModel> dataModelList);

        Tuple<string, List<D62ViewModel>> getD62(string reportDateStart, string reportDateEnd,
                                                 string referenceNbr, string bondNumber,
                                                 string basicPass,
                                                 string watchIND, string warningIND,
                                                 string chgInSpreadIND, string beforeHasChgInSpread);
        MSGReturnModel DownLoadD62Excel(string path, List<D62ViewModel> dbDatas);

        MSGReturnModel saveD62(string reportDate, string version);

        Tuple<string, D62ViewModel> getD62ByChg_In_Spread(D62ViewModel dataModel);

        MSGReturnModel modifyD62(D62ViewModel dataModel);

        #region D68
        Tuple<bool, List<D68ViewModel>> getD68All();

        Tuple<bool, List<D68ViewModel>> getD68(D68ViewModel dataModel);

        MSGReturnModel saveD68(string actionType, D68ViewModel dataModel);

        MSGReturnModel deleteD68(string ruleID);

        MSGReturnModel sendD68ToAudit(string ruleID, string auditor);

        MSGReturnModel D68Audit(string ruleID, string status);
        #endregion

        /// get Db D69 all data
        /// </summary>
        /// <returns></returns>
        Tuple<bool, List<D69ViewModel>> getD69All();

        /// <summary>
        /// get Db D69 data
        /// </summary>
        /// <param name="dataModel"></param>
        /// <returns></returns>
        Tuple<bool, List<D69ViewModel>> getD69(D69ViewModel dataModel);

        /// <summary>
        /// save D69 To Db
        /// </summary>
        /// <param name="dataModel">D69ViewModel</param>
        /// <returns></returns>
        MSGReturnModel saveD69(D69ViewModel dataModel);

        /// <summary>
        /// delete D69
        /// </summary>
        /// <param name="ruleID">規則編號</param>
        /// <returns></returns>
        MSGReturnModel deleteD69(string ruleID);

        /// <summary>
        /// send D69 to audit
        /// </summary>
        /// <param name="dataModelList"></param>
        /// <returns></returns>
        MSGReturnModel sendD69ToAudit(List<D69ViewModel> dataModelList);

        MSGReturnModel D69Audit(List<D69ViewModel> dataModelList);

        /// <summary>
        ///  下載 Excel
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="type">(D54)</param>
        /// <param name="path">下載位置</param>
        /// <param name="cache">cache 資料</param>
        /// <returns></returns>
        MSGReturnModel DownLoadExcel<T>(string type, string path, List<T> data);

        /// <summary>
        /// 下載 Excel (D63&D65 專用)
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="type"></param>
        /// <param name="path">D63,D65</param>
        /// <param name="path2">D64,D66</param>
        /// <param name="data"></param>
        /// <returns></returns>
        MSGReturnModel DownLoadExcel2<T>(string type, string path, string path2, List<T> data);

        Tuple<bool, List<D67ViewModel>> getD67(D67ViewModel dataModel);

        MSGReturnModel DownLoadD67Excel(string path, List<D67ViewModel> dbDatas);

        MSGReturnModel saveD67(string reportDate);

        /// <summary>
        /// get 減損作業狀態查詢 
        /// </summary>
        /// <param name="reportDate"></param>
        /// <returns></returns>
        List<D6CheckViewModel> getD6Check(string reportDate);

        #region 量化評估

        /// <summary>
        /// 量化評估
        /// </summary>
        /// <param name="reportDate"></param>
        /// <returns></returns>
        MSGReturnModel TransferD63(string reportDate);

        /// <summary>
        /// D63評估執行成功後 執行查詢(應該只會有版本1)
        /// </summary>
        /// <param name="reportDate"></param>
        /// <returns></returns>
        List<D63ViewModel> TransferD63Data(string reportDate);

        /// <summary>
        /// get D63 data
        /// </summary>
        /// <param name="reportDate">報導日</param>
        /// <param name="boundNumber">債券編號</param>
        /// <param name="indexFlag">評估狀態</param>
        /// <param name="assessmentSubKind">評估次分類</param>
        /// <param name="Send_to_AuditorFlag">是否提交複核</param>
        /// <param name="referenceNbr">帳戶編號</param>
        /// <returns></returns>
        List<D63ViewModel> getD63(DateTime reportDate, string boundNumber, Evaluation_Status_Type indexFlag, string assessmentSubKind, bool Send_to_AuditorFlag = false,string referenceNbr = "");

        /// <summary>
        /// get D63 Assessment_Result_Version
        /// </summary>
        /// <param name="Reference_Nbr"></param>
        /// <returns></returns>
        List<D63ViewModel> getD63History(string Reference_Nbr);

        /// <summary>
        /// get D64-量化評估結果檔
        /// </summary>
        /// <param name="referenceNbr"></param>
        /// <param name="Assessment_Result_Version"></param>
        /// <returns></returns>
        List<D64ViewModel> getD64(string referenceNbr, int Assessment_Result_Version);

        /// <summary>
        /// D63 複核
        /// </summary>
        /// <param name="action"></param>
        /// <param name="Auditor"></param>
        /// <param name="model"></param>
        /// <returns></returns>
        MSGReturnModel SaveD63(string action, string Auditor, D63ViewModel model);

        /// <summary>
        /// D63 複核系列動作
        /// </summary>
        /// <param name="Reference_Nbr"></param>
        /// <param name="Assessment_Result_Version"></param>
        /// <param name="status"></param>
        /// <returns></returns>
        MSGReturnModel UpdateD63(string Reference_Nbr,int Assessment_Result_Version, Evaluation_Status_Type status);
        #endregion 量化評估

        #region 質化評估

        /// <summary>
        /// 質化評估
        /// </summary>
        /// <param name="reportDate"></param>
        /// <returns></returns>
        MSGReturnModel TransferD65(string reportDate);

        /// <summary>
        /// D65評估執行成功後 執行查詢(應該只會有版本1)
        /// </summary>
        /// <param name="reportDate"></param>
        /// <returns></returns>
        List<D65ViewModel> TransferD65Data(string reportDate);

        /// <summary>
        /// get D65
        /// </summary>
        /// <param name="reportDate"></param>
        /// <param name="boundNumber"></param>
        /// <param name="referenceNbr"></param>
        /// <param name="indexFlag"></param>
        /// <param name="assessmentSubKind"></param>
        /// <param name="Send_to_AuditorFlag">是否提交複核</param>
        List<D65ViewModel> getD65(DateTime reportDate, string boundNumber, string referenceNbr, Evaluation_Status_Type indexFlag, string assessmentSubKind, bool Send_to_AuditorFlag = false);

        /// <summary>
        /// get D65 Assessment_Result_Version 
        /// </summary>
        /// <param name="Reference_Nbr"></param>
        /// <returns></returns>
        List<D65ViewModel> getD65History(string Reference_Nbr);

        /// <summary>
        /// get D66-質化評估結果檔
        /// </summary>
        /// <param name="referenceNbr">帳戶編號</param>
        /// <param name="Assessment_Result_Version">評估結果版本</param>
        /// <returns></returns>
        List<D66ViewModel> getD66(string referenceNbr, int Assessment_Result_Version);

        /// <summary>
        /// D65 複核推送
        /// </summary>
        /// <param name="Auditor"></param>
        /// <param name="D65Model">存D65 info ref_Nbr </param>
        /// <param name="D66Models">存D66 info 前端只送 pass_value & Check_Item_Code</param>
        /// <returns></returns>
        MSGReturnModel SaveD65(string Auditor, D65ViewModel D65Model, List<D66ViewModel> D66Models);

        /// <summary>
        /// D65 複核系列動作
        /// </summary>
        /// <param name="Reference_Nbr"></param>
        /// <param name="Assessment_Result_Version"></param>
        /// <param name="status"></param>
        /// <returns></returns>
        MSGReturnModel UpdateD65(string Reference_Nbr, int Assessment_Result_Version, Evaluation_Status_Type status);


        #endregion 質化評估

        #region 評估相關

        /// <summary>
        /// 刪除D64orD66附件檔案
        /// </summary>
        /// <param name="type"></param>
        /// <param name="Check_Reference"></param>
        /// <param name="fileName"></param>
        /// <returns></returns>
        MSGReturnModel DelQuantifyAndQualitativeFile(string type, string Check_Reference, string fileName);

        /// <summary>
        /// get QuantifyFile fileName
        /// </summary>
        /// <param name="Check_Reference"></param>
        /// <returns></returns>
        List<Bond_Quantitative_Result_File> getQuantifyFile(string Check_Reference);

        /// <summary>
        ///  get QualitativeFile fileName
        /// </summary>
        /// <param name="Check_Reference"></param>
        /// <returns></returns>
        List<Bond_Qualitative_Assessment_Result_File> getQualitativeFile(string Check_Reference);

        /// <summary>
        /// save QualitativeFile
        /// </summary>
        /// <param name="Check_Reference"></param>
        /// <param name="FileName"></param>
        /// <param name="type"></param>
        /// <param name="File_path"></param>
        /// <returns></returns>
        MSGReturnModel SaveQualitativeFile(string Check_Reference, string FileName, Table_Type type, string File_path);

        /// <summary>
        /// get Assessment_Result , Memo
        /// </summary>
        /// <param name="ReportDate"></param>
        /// <param name="Reference_Nbr"></param>
        /// <param name="Version"></param>
        /// <param name="Assessment_Result_Version"></param>
        /// <param name="type"></param>
        /// <param name="table"></param>
        /// <returns></returns>
        string GetTextArea(DateTime ReportDate, string Reference_Nbr, int Version, int Assessment_Result_Version, string Check_Item_Code, string type, Table_Type table);

        /// <summary>
        /// Save Assessment_Result , Memo
        /// </summary>
        /// <param name="value"></param>
        /// <param name="ReportDate"></param>
        /// <param name="Reference_Nbr"></param>
        /// <param name="Vresion"></param>
        /// <param name="Assessment_Result_Version"></param>
        /// <param name="type"></param>
        /// <param name="table"></param>
        /// <returns></returns>
        MSGReturnModel SaveTextArea(string value, DateTime ReportDate, string Reference_Nbr, int Vresion, int Assessment_Result_Version, string Check_Item_Code, string type, Table_Type table);

        /// <summary>
        /// get Assessment Sub Kind
        /// </summary>
        /// <param name="type"></param>
        /// <returns></returns>
        List<SelectOption> getAssessmentSubKind(AssessmentKind_Type type);

        /// <summary>
        /// get 減損試算結果
        /// </summary>
        /// <param name="Reference_Nbr"></param>
        /// <returns></returns>
        List<ChgInSpreadViewModel> getChgInSpread(string Reference_Nbr);

        /// <summary>
        /// get Assessment_Stage
        /// </summary>
        /// <param name="Assessment_Kind"></param>
        /// <param name="Assessment_Sub_Kind"></param>
        /// <returns></returns>
        List<string> getAssessmentStage(string Assessment_Kind, string Assessment_Sub_Kind);

        /// <summary>
        /// get Check Item Code
        /// </summary>
        /// <param name="Assessment_Stage"></param>
        /// <param name="Assessment_Kind"></param>
        /// <param name="Assessment_Sub_Kind"></param>
        /// <returns></returns>
        List<Tuple<string, string>> getCheckItemCode(string Assessment_Stage, string Assessment_Kind, string Assessment_Sub_Kind);
        #endregion 評估相關

        #region Extra_Case

        int GetA41Version(string reportDate);

        MSGReturnModel DeleteD65ByExtraCase(string referenceNbr, int version);

        MSGReturnModel InsertD65ByExtraCase(string reportDate, int version, string bondNumber, string Lots, string portfolio_Name);

        #endregion

        /// <summary>
        /// get EL FLOW_ID
        /// </summary>
        /// <param name="reportDate"></param>
        /// <returns></returns>
        List<string> getFlowIDs(string reportDate);

        /// <summary>
        /// 查詢 ReEL
        /// </summary>
        /// <param name="reportDate"></param>
        /// <returns></returns>
        List<ReELViewModel> getReEL(string reportDate,string Group_Product_Code);

        /// <summary>
        /// 減損報表資料刪除作業
        /// </summary>
        /// <param name="reportDate"></param>
        /// <param name="flowId"></param>
        /// <param name="msg"></param>
        /// <returns></returns>
        MSGReturnModel ReEL(string reportDate, string flowId, string msg);
    }
}