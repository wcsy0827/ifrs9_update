using System;
using System.Collections.Generic;
using System.IO;
using Transfer.Utility;
using Transfer.ViewModels;
using static Transfer.Enum.Ref;

namespace Transfer.Models.Interface
{
    public interface IC0Repository
    {
        List<string> getC01Version(string reportDate, string productCode);

        List<string> GetC01LogData(string tableTypes, string debt);

        Tuple<bool, List<C01ViewModel>> getC01(string reportDate, string productCode, string version);

        List<C01ViewModel> getC01Excel(string pathType, string path, string version);

        MSGReturnModel saveC01(string country, string version, List<C01ViewModel> dataModel);

        List<SelectOption> getProduct(GroupProductCode type);

        List<string> getC07Version(string debtType, string productCode, string reportDate);

        /// <summary>
        /// get Db C07 data
        /// </summary>
        /// <returns></returns>
        Tuple<bool, List<C07ViewModel>> getC07(string debtType, string groupProductCode, string productCode, string reportDate, string version);

        /// <summary>
        /// 下載 Excel
        /// </summary>
        /// <param name="type">(C07Mortgage.C07Bond)</param>
        /// <param name="path">下載位置</param>
        /// <param name="dbDatas"></param>
        /// <returns></returns>
        MSGReturnModel DownLoadExcel(string type, string path, List<C07ViewModel> dbDatas);

        /// <summary>
        /// 下載C10Excel
        /// </summary>
        /// <param name="type"></param>
        /// <param name="path">下載位置</param>
        /// <param name="data"></param>
        /// <returns></returns>
        MSGReturnModel DownLoadExcelC10(string type, string path, List<C10DetailViewModel> data); //190625 PGE需求新增


        List<string> getA41AdvancedAssessment_Sub_Kind();

        Tuple<string, List<C07AdvancedViewModel>> getC07Advanced(string groupProductCode, string productCode, string reportDate, string version, string assessmentSubKind, string watchIND,bool isReport=false);

        MSGReturnModel DownloadC07AdvancedExcel(string type, string path, List<C07AdvancedViewModel> dbDatas);

        Tuple<bool, List<C07AdvancedSumViewModel>> getC07AdvancedSum(string reportDate, string version, string groupProductCode, string groupProductName, string referenceNbr,string assessmentSubKind, string watchIND,string productCode);
        //Tuple<bool, List<SummaryViewModel>> getSum(string reportDate, string version, string groupProductCode, string groupProductName, string referenceNbr, string assessmentSubKind, string watchIND, string productCode);

        //SummaryViewModel
        MSGReturnModel DownloadC07AdvancedSumExcel(string type, string path, List<C07AdvancedSumViewModel> dbDatas);

        /// <summary>
        /// Get C04 View Data 
        /// </summary>
        /// <param name="ds">起始季別</param>
        /// <param name="de">截止季別</param>
        /// <param name="lastFlag">僅顯示最近更新資料</param>
        /// <returns></returns>
        List<C04ViewModel> GetC04(string ds, string de, bool lastFlag = false);

        /// <summary>
        /// Get C04 數列資料的時間
        /// </summary>
        /// <returns></returns>
        List<string> GetC04YearQuartly();

        List<C04ProViewModel> GetC04Pro();

        /// <summary>
        /// 下載 Excel
        /// </summary>
        /// <param name="type">(C04_1)</param>
        /// <param name="path">下載位置</param>
        /// <param name="cache">cache 資料</param>
        /// <returns></returns>
        MSGReturnModel DownLoadExcel<T>(string type, string path, List<T> data);

        /// <summary>
        /// 
        /// </summary>
        /// <param name="from"></param>
        /// <param name="to"></param>
        /// <param name="data"></param>
        /// <returns></returns>
        List<System.Dynamic.ExpandoObject> C04Transfer(string from, string to, List<C04ProViewModel> data);

        Tuple<string, List<C10ViewModel>> getExcel(string pathType, Stream stream, DateTime reportDate);

        MSGReturnModel saveC10(List<C10ViewModel> dataModel, string reportDate);

        List<C10DetailViewModel> GetC10(DateTime ReportDate, int Version);
    }
}