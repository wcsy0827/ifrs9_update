using System;
using System.Collections.Generic;
using System.IO;
using Transfer.Utility;
using Transfer.ViewModels;
using static Transfer.Enum.Ref;

namespace Transfer.Models.Interface
{
    public interface IA9Repository
    {
        Tuple<bool, List<A94ViewModel>> getA94All();

        Tuple<bool, List<A94ViewModel>> getA94(A94ViewModel dataModel);

        MSGReturnModel saveA94(string actionType, A94ViewModel dataModel);

        MSGReturnModel deleteA94(string country);

        #region A95_1 (產業別對應Ticker補登檔)
        /// <summary>
        /// Get A95_1 產業別對應Ticker補登檔 
        /// </summary>
        /// <param name="bondNumber"></param>
        /// <returns></returns>
        List<A95_1ViewModel> getA95_1(string bondNumber);

        /// <summary>
        /// insert A95_1 (產業別對應Ticker補登檔)
        /// </summary>
        /// <param name="datas"></param>
        /// <returns></returns>
        MSGReturnModel insertA95_1(List<A95_1ViewModel> datas);

        /// <summary>
        /// save A95_1 (產業別對應Ticker補登檔)
        /// </summary>
        /// <param name="data"></param>
        /// <param name="type"></param>
        /// <returns></returns>
        MSGReturnModel saveA95_1(A95_1ViewModel data, Action_Type type);

        /// <summary>
        /// Excel 資料轉成 A95_1ViewModel
        /// </summary>
        /// <param name="pathType"></param>
        /// <param name="stream"></param>
        /// <returns></returns>
        Tuple<string, List<A95_1ViewModel>> getA95_1Excel(string pathType, Stream stream);
        #endregion

        /// <summary>
        ///  get A96 資料
        /// </summary>
        /// <param name="dt"></param>
        /// <param name="version"></param>
        /// <returns></returns>
        List<A96ViewModel> getA96(DateTime dt, int version);

        /// <summary>
        /// get A96 最後交易日 資料
        /// </summary>
        /// <param name="dt"></param>
        /// <returns></returns>
        List<A96TradeViewModel> getA96Trade(DateTime? dt);

        /// <summary>
        /// get A96 version
        /// </summary>
        /// <param name="dt"></param>
        /// <returns></returns>
        List<string> getA96Version(DateTime dt);

        /// <summary>
        /// save A96 信用利差資料
        /// </summary>
        /// <param name="datas"></param>
        /// <returns></returns>
        MSGReturnModel saveA96(List<A96ViewModel> datas);

        /// <summary>
        /// 新增,刪除,修改 A96 最後交易日
        /// </summary>
        /// <param name="data"></param>
        /// <param name="type"></param>
        /// <returns></returns>
        MSGReturnModel saveA96Trade(A96TradeViewModel data, Action_Type type);

        /// <summary>
        /// 下載 Excel
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="type"></param>
        /// <param name="path"></param>
        /// <param name="data"></param>
        /// <returns></returns>
        MSGReturnModel DownLoadExcel<T>(Excel_DownloadName type, string path, List<T> data);

        /// <summary>
        /// Excel 資料轉成 A96ViewModel
        /// </summary>
        /// <param name="pathType"></param>
        /// <param name="stream"></param>
        /// <returns></returns>
        Tuple<string, List<A96ViewModel>> getA96Excel(string pathType, Stream stream);
    }
}