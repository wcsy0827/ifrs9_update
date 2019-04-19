using System;
using System.Collections.Generic;
using System.IO;
using Transfer.Utility;
using Transfer.ViewModels;

namespace Transfer.Models.Interface
{
    public interface ID7Repository
    {
        Tuple<bool, List<D70ViewModel>> getD70All(string type);

        Tuple<bool, List<D70ViewModel>> getD70(string type, D70ViewModel dataModel);

        MSGReturnModel saveD70(string type, string actionType, D70ViewModel dataModel);

        MSGReturnModel deleteD70(string type, string ruleID);

        MSGReturnModel sendD70ToAudit(string type, List<D70ViewModel> dataModelList);

        MSGReturnModel D70Audit(string type, List<D70ViewModel> dataModelList);

        /// <summary>
        /// Get D72
        /// </summary>
        /// <returns></returns>
        Tuple<string, List<D72ViewModel>> GetD72();

        /// <summary>
        /// Save D72
        /// </summary>
        /// <param name="datas"></param>
        /// <returns></returns>
        MSGReturnModel SaveD72(List<D72ViewModel> datas);

        /// <summary>
        /// Excel資料 轉 D72ViewModel
        /// </summary>
        /// <param name="pathType"></param>
        /// <param name="stream"></param>
        /// <returns></returns>
        Tuple<string, List<D72ViewModel>> GetD72Excel(string pathType, Stream stream);

        /// <summary>
        /// Get D73 Version
        /// </summary>
        /// <param name="date"></param>
        /// <returns></returns>
        List<string> GetD73Veriosn(DateTime? date);

        /// <summary>
        /// 檢視 D73 檔案內容
        /// </summary>
        /// <param name="Id"></param>
        /// <returns></returns>
        string GetD73FileLog(int Id);

        /// <summary>
        /// get D73 (排程查核結果紀錄檔)
        /// </summary>
        /// <param name="date">基準日</param>
        /// <param name="ver">版本</param>
        /// <returns></returns>
        List<D73ViewModel> GetD73(DateTime? date, int? ver);

        /// <summary>
        /// 刪除 D73 Log檔案
        /// </summary>
        /// <param name="ID"></param>
        /// <returns></returns>
        MSGReturnModel DelD73(string ID);

        /// <summary>
        /// 抓取啟用的主標尺對應檔年度
        /// </summary>
        /// <returns></returns>
        string GetA51Year();

        /// <summary>
        /// 找尋設定的 通知名稱
        /// </summary>
        /// <param name="type"></param>
        /// <returns></returns>
        List<SelectOption> GetNoticeName(string type);

        /// <summary>
        /// Get D74
        /// </summary>
        /// <param name="noticeName"></param>
        /// <returns></returns>
        List<D74ViewModel> GetD74(string noticeID);

        /// <summary>
        ///  Get D74_1 
        /// </summary>
        /// <param name="noticeID"></param>
        /// <returns></returns>
        List<D74_1ViewModel> GetD74_1(string noticeID);

        /// <summary>
        /// 新增通知設定檔
        /// </summary>
        /// <param name="data"></param>
        /// <returns></returns>
        MSGReturnModel AddD74(D74ViewModel data);

        /// <summary>
        /// 修改通知設定檔
        /// </summary>
        /// <param name="data"></param>
        /// <returns></returns>
        MSGReturnModel SaveD74(D74ViewModel data);

        /// <summary>
        /// 新增郵件設定檔
        /// </summary>
        /// <param name="data"></param>
        /// <returns></returns>
        MSGReturnModel AddD74_1( D74_1ViewModel data);

        /// <summary>
        /// 修改郵件設定檔
        /// </summary>
        /// <param name="data"></param>
        /// <returns></returns>
        MSGReturnModel SaveD74_1(D74_1ViewModel data);

        /// <summary>
        /// 刪除郵件設定檔
        /// </summary>
        /// <param name="data"></param>
        /// <returns></returns>
        MSGReturnModel DeleD74_1(D74_1ViewModel data);
    }
}