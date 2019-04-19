using System;
using System.Collections.Generic;
using System.IO;
using Transfer.Utility;
using Transfer.ViewModels;

namespace Transfer.Models.Interface
{
    public interface IA4Repository
    {
        /// <summary>
        /// get Db A41 data
        /// </summary>
        /// <returns></returns>
        List<A41ViewModel> GetA41(DateTime ReportDate, DateTime OriginationDate, int Version, string BondNumber);

        /// <summary>
        /// get Db A42 data
        /// </summary>
        /// <param name="reportDate">評估基準日/報導日</param>
        /// <returns></returns>
        Tuple<bool, List<A42ViewModel>> getA42(string reportDate);

        /// <summary>
        /// 判斷 A42 是否重複上傳
        /// </summary>
        /// <param name="reportDate">reportDate</param>
        /// <returns></returns>
        MSGReturnModel checkA42Duplicate(string reportDate);

        List<A42ViewModel> getA42Excel(string pathType, string path, string processingDate, string reportDate);

        /// <summary>
        /// save A42 To Db
        /// </summary>
        /// <param name="dataModel">A42ViewModel</param>
        /// <returns></returns>
        MSGReturnModel saveA42(List<A42ViewModel> dataModel);

        MSGReturnModel DownLoadA42Excel(string path, List<A42ViewModel> dbDatas);

        /// <summary>
        /// get Db A45 data
        /// </summary>
        /// <param name="bloombergTicker">機構簡稱</param>
        /// <param name="processingDate">資料處理日期</param>
        /// <returns></returns>
        Tuple<bool, List<A45ViewModel>> getA45(string bloombergTicker, string processingDate);

        /// <summary>
        /// get Db A46 data
        /// </summary>
        /// <param name="searchAll"></param>
        /// <returns></returns>
        List<A46ViewModel> GetA46(bool searchAll);

        /// <summary>
        /// get Db A47 data
        /// </summary>
        /// <param name="searchAll"></param>
        /// <returns></returns>
        List<A47ViewModel> GetA47(bool searchAll);

        /// <summary>
        /// get Db A48 data
        /// </summary>
        /// <param name="searchAll"></param>
        /// <returns></returns>
        List<A48ViewModel> GetA48(bool searchAll);

        /// <summary>
        /// get A95 data
        /// </summary>
        /// <param name="reportDate"></param>
        /// <param name="verison"></param>
        /// <param name="bondType">債券種類</param>
        /// <param name="ASK">產業別</param>
        /// <returns></returns>
        Tuple<bool, List<A95ViewModel>> GetA95(DateTime reportDate, int verison, string bondType, string ASK);

        /// <summary>
        /// Excel資料 轉 A45ViewModel
        /// </summary>
        /// <param name="pathType"></param>
        /// <param name="stream"></param>
        /// <returns></returns>
        List<A45ViewModel> getA45Excel(string pathType, Stream stream);

        /// <summary>
        /// Excel資料 轉 Exhibit29Model
        /// </summary>
        /// <param name="pathType"></param>
        /// <param name="stream"></param>
        /// <param name="reportDate"></param>
        /// <param name="version"></param>
        /// <returns></returns>
        Tuple<string,List<A41ViewModel>> getExcel(string pathType, Stream stream, DateTime reportDate, int version);

        /// <summary>
        /// Excel資料 轉 A95ViewModel
        /// </summary>
        /// <param name="pathType"></param>
        /// <param name="path"></param>
        /// <returns></returns>
        List<A95ViewModel> getA95Excel(string pathType, string path);

        /// <summary>
        /// 抓取指定的 log 資料
        /// </summary>
        /// <param name="tableTypes">"B01","C01"...</param>
        /// <returns></returns>
        List<string> GetLogData(List<string> tableTypes, string debt);

        /// <summary>
        /// save A41 To Db
        /// </summary>
        /// <param name="dataModel">A41ViewModel</param>
        /// <returns></returns>
        MSGReturnModel saveA41(List<A41ViewModel> dataModel, string reportDate, string version);

        /// <summary>
        /// save A45 To Db
        /// </summary>
        /// <param name="dataModel">A45ViewModel</param>
        /// <returns></returns>
        MSGReturnModel saveA45(List<A45ViewModel> dataModel);

        /// <summary>
        /// Excel資料 轉 A46ViewModel
        /// </summary>
        /// <param name="pathType"></param>
        /// <param name="stream"></param>
        /// <returns></returns>
        Tuple<string, List<A46ViewModel>> getA46Excel(string pathType, Stream stream);

        /// <summary>
        /// save A46 To Db
        /// </summary>
        /// <param name="datas">A46ViewModel</param>
        /// <returns></returns>
        MSGReturnModel saveA46(List<A46ViewModel> datas);

        /// <summary>
        /// Excel資料 轉 A47ViewModel
        /// </summary>
        /// <param name="pathType"></param>
        /// <param name="stream"></param>
        /// <returns></returns>
        Tuple<string, List<A47ViewModel>> getA47Excel(string pathType, Stream stream);

        /// <summary>
        /// save A47 To Db
        /// </summary>
        /// <param name="datas">A47ViewModel</param>
        /// <returns></returns>
        MSGReturnModel saveA47(List<A47ViewModel> datas);

        /// <summary>
        /// Excel資料 轉 A48ViewModel
        /// </summary>
        /// <param name="pathType"></param>
        /// <param name="stream"></param>
        /// <returns></returns>
        Tuple<string, List<A48ViewModel>> getA48Excel(string pathType, Stream stream);

        /// <summary>
        /// save A48 To Db
        /// </summary>
        /// <param name="datas">A48ViewModel</param>
        /// <returns></returns>
        MSGReturnModel saveA48(List<A48ViewModel> datas);

        /// <summary>
        /// save A95 To Db
        /// </summary>
        /// <param name="dataModel"></param>
        /// <returns></returns>
        MSGReturnModel saveA95(List<A95ViewModel> dataModel);

        /// <summary>
        /// save B01 to DB
        /// </summary>
        /// <param name="version">version</param>
        /// <param name="date">Report_Date</param>
        /// <returns></returns>
        MSGReturnModel saveB01(int version, DateTime date, string type);

        /// <summary>
        /// save C01 to DB
        /// </summary>
        /// <param name="version">version</param>
        /// <param name="date">Report_Date</param>
        /// <returns></returns>
        MSGReturnModel saveC01(int version, DateTime date, string type);

        /// <summary>
        /// save C02 to DB
        /// </summary>
        /// <param name="version">version</param>
        /// <param name="date">Report_Date</param>
        /// <param name="type">M = 房貸</param>
        /// <returns></returns>
        MSGReturnModel saveC02(int version, DateTime date, string type);

        /// <summary>
        /// DownLoadExcel
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="type"></param>
        /// <param name="path"></param>
        /// <param name="data"></param>
        /// <returns></returns>
        MSGReturnModel DownLoadExcel<T>(string type, string path, List<T> data);

        Tuple<bool, List<A44ViewModel>> getA44All();
        Tuple<bool, List<A44ViewModel>> getA44(A44ViewModel dataModel);
        MSGReturnModel saveA44(string actionType, A44ViewModel dataModel);
        MSGReturnModel deleteA44(string bondNumberNew, string lotsNew, string portfolioNameNew);

        Tuple<bool, List<A49ViewModel>> getA49(A49ViewModel dataModel);
        List<A49ViewModel> getA49Excel(string pathType, Stream stream);
        MSGReturnModel saveA49(List<A49ViewModel> dataModel, string reportDate);
    }
}