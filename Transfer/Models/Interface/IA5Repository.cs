using System;
using System.Collections.Generic;
using Transfer.Utility;
using Transfer.ViewModels;
using static Transfer.Enum.Ref;

namespace Transfer.Models.Interface
{
    public interface IA5Repository
    {
        Tuple<bool, List<A51ViewModel>> getA51(Audit_Type type,string dataYear, string rating = null, string pdGrade = null, string ratingAdjust = null, 
                                               string gradeAdjust = null, string moodysPD = null);

        bool getA53(DateTime date);

        Tuple<bool, List<A52ViewModel>> getA52All();

        /// <summary>
        /// get A52 
        /// </summary>
        /// <param name="ratingOrg">RatingOrg</param>
        /// <param name="pdGrade">PdGrade</param>
        /// <param name="rating">Rating</param>
        /// <param name="IsActive">是否有效(All:全部,Y:生效,N:失效)</param>
        /// <param name="Status">複核結果(All:全部)</param>
        /// <returns></returns>
        Tuple<bool, List<A52ViewModel>> getA52(string ratingOrg = "All", string pdGrade = "All", string rating = "All", string IsActive = "All", string Status = "All");

        MSGReturnModel saveA52(string actionType, A52ViewModel dataModel);

        MSGReturnModel deleteA52(A52ViewModel dataModel);

        /// <summary>
        /// Get A56
        /// </summary>
        /// <param name="IsActive">是否生效</param>
        /// <param name="Replace_Object">特殊字元</param>
        /// <returns></returns>
        List<A56ViewModel> GetA56(string IsActive, string Replace_Object);

        /// <summary>
        /// Get A56 Replace_Object
        /// </summary>
        /// <returns></returns>
        List<string> GetA56_Replace_Object();

        /// <summary>
        /// get A57 Data
        /// </summary>
        /// <param name="datepicker"></param>
        /// <param name="version"></param>
        /// <param name="df"></param>
        /// <param name="dt"></param>
        /// <param name="SMF"></param>
        /// <param name="stype"></param>
        /// <param name="bondNumber"></param>
        /// <param name="issuer"></param>
        /// <param name="status"></param>
        /// <returns></returns>
        List<A57ViewModel> GetA57(DateTime datepicker, int version, DateTime? df, DateTime? dt, string SMF, string stype, string bondNumber, string issuer, Rating_Status status);

        /// <summary>
        /// get A58 Data
        /// </summary>
        /// <param name="datepicker">ReportDate</param>
        /// <param name="sType">Rating_Type</param>
        /// <param name="from">Origination_Date start</param>
        /// <param name="to">Origination_Date to</param>
        /// <param name="bondNumber">bondNumber</param>
        /// <param name="version">version</param>
        /// <param name="search">全部or缺漏</param>
        /// <param name="portfolio">portfolio</param>
        /// <returns></returns>
        Tuple<bool, List<A58ViewModel>> GetA58(string datepicker, string sType, string from, string to, string bondNumber, string version, string search,string portfolio);

        /// <summary>
        ///  Excel 資料轉成 A59ViewModel
        /// </summary>
        /// <param name="pathType">Excel 副檔名</param>
        /// <param name="path">檔案路徑</param>
        /// <param name="action">目前動作</param>
        /// <returns></returns>
        Tuple<string, List<A59ViewModel>> getA59Excel(string pathType, string path, string action);


        //190222 接取寶碩DB信評補進A59
        /// <summary>
        ///  取得自動補檔後A59
        /// </summary>
        /// <param name="datamodel">缺信評A58ViewModel</param>
        /// <returns></returns>
        Tuple<bool, List<A59ViewModel>> getA59(List<A58ViewModel> datamodel, string reportdate);

        //190222 A59轉入A57A58結果寫入Transfer_CheckTable Log
        /// <summary>
        ///  SaveA59、A57、A58結果寫入Log
        /// </summary>
        ///<param name="result_autotrasfer">SaveA59執行結果</param>
        ///<param name="reportdate">報導日</param>
        ///<param name="starttime">執行開始時間</param>
        ///<param name="version">轉入版本</param>
        /// <returns></returns>
        void SaveA59TransLog(MSGReturnModel result_autotrasfer, string reportdate, DateTime starttime, int version);

        //190222 將自動補檔後的A59檔案存至指定位置
        /// <summary>
        ///  取得自動補檔後A59檔案
        /// </summary>
        ///<param name="type">存檔檔案名稱</param>
        ///<param name="path">存檔路徑</param>
        /// <param name="datamodel">補完寶碩信評A59ViewModel</param>
        /// <returns></returns>
        MSGReturnModel SaveA59Excel(string type, string path, List<A59ViewModel> datamodel);


        //190222 自動補檔後檢核
        /// <summary>
        ///  自動補檔後檢核
        /// </summary>
        /// <param name="reportdate">報導日</param>
        /// <param name="version">版本</param>
        /// <returns></returns>
        void GetA58TransferCheck(DateTime reportdate, int version);


        /// <summary>
        /// get 轉檔紀錄Table 資料
        /// </summary>
        /// <param name="fileNames"></param>
        /// <returns></returns>
        List<CheckTableViewModel> getCheckTable(List<string> fileNames);

        /// <summary>
        /// get SMF
        /// </summary>
        /// <param name="date"></param>
        /// <param name="version"></param>
        /// <returns></returns>
        List<string> getSMF(DateTime date, int version);

        /// <summary>
        ///  下載 Excel
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="type">(A59)</param>
        /// <param name="path">下載位置</param>
        /// <param name="cache">cache 資料</param>
        /// <returns></returns>
        MSGReturnModel DownLoadExcel<T>(string type, string path, List<T> data, bool flag = false);

        /// <summary>
        /// Save A56
        /// </summary>
        /// <param name="data">新增 or 修改(限定只能設定失效)</param>
        /// <param name="IsActive"></param>
        /// <returns></returns>
        MSGReturnModel saveA56(A56ViewModel data, bool IsActive);

        /// <summary>
        /// save A59 To Db
        /// </summary>
        /// <param name="dataModel">A59ViewModel</param>
        /// <returns></returns>
        MSGReturnModel saveA59(List<A59ViewModel> dataModel);

        /// <summary>
        /// 手動轉換 A57 & A58
        /// </summary>
        /// <param name="date"></param>
        /// <param name="version"></param>
        /// <param name="complement">抓取上一版資料(評估日)</param>
        /// <param name="deleteFlag">重新執行最新版</param>
        /// <param name="A59Flag">A59補登Flag</param>
        /// <returns></returns>
        MSGReturnModel saveA57A58(DateTime date, int version, string complement, bool deleteFlag, bool A59Flag = false);

        /// <summary>
        /// 檢查是不是該基準日最後一版
        /// </summary>
        /// <param name="date"></param>
        /// <param name="version"></param>
        /// <returns></returns>
        MSGReturnModel checkVersion(DateTime date, int version);

        Tuple<List<SelectOption>, List<SelectOption>, List<SelectOption>> GetA52SearchData(string ratingOrg, string IsActive, string pdGrade, string rating);

        /// <summary>
        /// 查詢 A52(信評主標尺對應檔_其他) by Audit_date for selectoption
        /// </summary>
        /// <param name="Audit_date"></param>
        /// <returns></returns>
        List<SelectOption> GetA52Auditdate(string Audit_date);

        /// <summary>
        /// 查詢 A52(信評主標尺對應檔_其他) by Audit_date 
        /// </summary>
        /// <param name="Audit_date"></param>
        /// <returns></returns>
        List<A52ViewModel> GetA52byAuditdate(string Audit_date);

        /// <summary>
        /// A52(信評主標尺對應檔_其他) 複核
        /// </summary>
        /// <param name="Id"></param>
        /// <param name="Status"></param>
        /// <param name="Auditor_Reply"></param>
        /// <returns></returns>
        MSGReturnModel A52Audit(int? Id, string Status, string Auditor_Reply);
    }
}