using System;
using System.Collections.Generic;
using Transfer.Utility;
using static Transfer.Enum.Ref;

namespace Transfer.Models.Interface
{
    public interface ICommon
    {
        /// <summary>
        /// 判斷轉檔紀錄是否有存在
        /// </summary>
        /// <param name="fileNames">目前檔案名稱</param>
        /// <param name="checkName">要檢查的檔案名稱</param>
        /// <param name="reportDate">基準日</param>
        /// <param name="version">版本</param>
        /// /// <returns></returns>
        bool checkTransferCheck(
            string fileName,
            string checkName,
            DateTime reportDate,
            int version);

        /// <summary>
        /// Log資料存到Sql(IFRS9_Log)
        /// </summary>
        /// <param name="tableType">table簡寫</param>
        /// <param name="tableName">table名</param>
        /// <param name="fileName">檔案名</param>
        /// <param name="programName">專案名</param>
        /// <param name="falg">成功失敗</param>
        /// <param name="deptType">B:債券 M:房貸 (共用同一table時做區分)</param>
        /// <param name="start">開始時間</param>
        /// <param name="end">結束時間</param>
        /// <param name="userAccount">執行帳號</param>
        /// <param name="version">版本</param>
        /// <param name="reportDate">基準日</param>
        /// <returns>回傳成功或失敗</returns>
        bool saveLog(
            Table_Type table,
            string fileName,
            string programName,
            bool falg,
            string deptType,
            DateTime start,
            DateTime end,
            string userAccount,
            int version = 1,
            Nullable<DateTime> date = null);

        /// <summary>
        /// 轉檔紀錄存到Sql(Transfer_CheckTable)
        /// </summary>
        /// <param name="fileName">檔案名稱 A41,A42...</param>
        /// <param name="falg">成功失敗</param>
        /// <param name="reportDate">基準日</param>
        /// <param name="version">版本</param>
        /// <param name="start">轉檔開始時間</param>
        /// <param name="end">轉檔結束時間</param>
        /// <param name="Msg">訊息</param>
        /// <returns>回傳成功或失敗</returns>
        bool saveTransferCheck(
             string fileName,
             bool falg,
             DateTime reportDate,
             int version,
             DateTime start,
             DateTime end,
             string Msg = null);

        /// <summary>
        /// reportDate 選擇後 查詢目前有的版本
        /// </summary>
        /// <param name="date"></param>
        /// <param name="table"></param>
        /// <returns></returns>
        List<string> getVersion(DateTime date, string tableName);

        /// <summary>
        /// get 複核者 or 評估者
        /// </summary>
        /// <param name="productCode"></param>
        /// <param name="tableId"></param>
        /// <param name="type"></param>
        /// <returns></returns>
        List<IFRS9_User> getAssessmentInfo(string productCode, string tableId, SetAssessmentType type);

        /// <summary>
        /// get Debt SelectOption
        /// </summary>
        /// <param name="account"></param>
        /// <returns></returns>
        List<SelectOption> getDebtSelectOption(string account);

        /// <summary>
        /// get user debt
        /// </summary>
        /// <param name="account"></param>
        /// <returns></returns>
        string getUserDebt(string account);
    }
}