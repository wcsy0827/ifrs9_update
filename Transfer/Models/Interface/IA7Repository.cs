using System;
using System.Collections.Generic;
using System.IO;
using Transfer.Utility;
using Transfer.ViewModels;

namespace Transfer.Models.Interface
{
    public interface IA7Repository
    {
        /// <summary>
        /// 下載 Excel
        /// </summary>
        /// <param name="type">(A72.A73)</param>
        /// <param name="path">下載位置</param>
        /// <returns></returns>
        MSGReturnModel DownLoadExcel(string type, string path);

        /// <summary>
        /// Get A51 Data
        /// </summary>
        /// <returns></returns>
        Tuple<bool, List<A51ViewModel>> GetA51(string year);

        /// <summary>
        /// Get A71 Data
        /// </summary>
        /// <returns></returns>
        Tuple<bool, List<Moody_Tm_YYYY>> GetA71();

        /// <summary>
        /// Get A71 Data
        /// </summary>
        /// <returns></returns>
        Tuple<bool, List<object>> GetA72();

        /// <summary>
        /// Get A73 Data
        /// </summary>
        /// <returns></returns>
        Tuple<bool, List<object>> GetA73();

        /// <summary>
        /// Excel資料 轉 Exhibit29Model
        /// </summary>
        /// <param name="pathType"></param>
        /// <param name="stream"></param>
        /// <returns></returns>
        Tuple<int, List<Exhibit29Model>> getExcel(string pathType, Stream stream);

        /// <summary>
        /// get A51 year
        /// </summary>
        /// <returns></returns>
        List<string> GetA51SearchYear();

        /// <summary>
        ///  save A51 To Db
        /// </summary>
        /// <param name="Year"></param>
        /// <returns></returns>
        MSGReturnModel saveA51(int Year);

        /// <summary>
        /// save A71 To Db
        /// </summary>
        /// <param name="dataModel">Exhibit29Model</param>
        /// <returns></returns>
        MSGReturnModel saveA71(List<Exhibit29Model> dataModel);

        /// <summary>
        /// save A72 To Db
        /// </summary>
        /// <param name="dataModel"></param>
        /// <returns></returns>
        MSGReturnModel saveA72();

        /// <summary>
        /// save A73 To Db
        /// </summary>
        /// <param name="dataModel"></param>
        /// <returns></returns>
        MSGReturnModel saveA73();

        /// <summary>
        /// 選擇A51複核者
        /// </summary>
        /// <param name="year">年份</param>
        /// <param name="Auditor">複核者</param>
        /// <returns></returns>
        MSGReturnModel AssessmentA51(string year, string Auditor);

        /// <summary>
        /// A51複核者只能變更 啟用 or 暫不啟用
        /// </summary>
        /// <param name="type">啟用 or 暫不啟用</param>
        /// <param name="year">年份</param>
        /// <param name="Auditor_Reply">複核訊息</param>
        /// <returns></returns>
        MSGReturnModel AuditA51(Enum.Ref.Audit_Type type, string year, string Auditor_Reply);

        /// <summary>
        /// 檢查可否上傳 A51
        /// </summary>
        /// <param name="year"></param>
        /// <returns></returns>
        bool getA51SaveFlag(int year);
    }
}