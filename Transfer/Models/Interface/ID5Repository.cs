using System;
using System.Collections.Generic;
using System.IO;
using Transfer.Utility;
using Transfer.ViewModels;

namespace Transfer.Models.Interface
{
    public interface ID5Repository
    {
        /// <summary>
        /// 查詢D53
        /// </summary>
        /// <param name="model"></param>
        /// <returns></returns>
        Tuple<string, List<D53ViewModel>> GetD53();

        /// <summary>
        /// Save D53
        /// </summary>
        /// <param name="datas"></param>
        /// <returns></returns>
        MSGReturnModel SaveD53(List<D53ViewModel> datas);

        /// <summary>
        /// Excel 資料轉成 D53ViewModel
        /// </summary>
        /// <param name="pathType"></param>
        /// <param name="stream"></param>
        /// <returns></returns>
        Tuple<string, List<D53ViewModel>> GetD53Excel(string pathType, Stream stream);

        /// <summary>
        /// D54查詢預計調整資料
        /// </summary>
        /// <param name="dt"></param>
        /// <returns></returns>
        List<D54ViewModel> getD54InsertSearch(string dt);

        /// <summary>
        /// 減損階段確認
        /// </summary>
        /// <param name="dt"></param>
        /// <returns></returns>
        MSGReturnModel SaveD54(string dt);

        /// <summary>
        /// get D54 Group_Product_Code
        /// </summary>
        /// <returns></returns>
        List<SelectOption> getD54GroupProductCode();

        /// <summary>
        /// get D54
        /// </summary>
        /// <param name="reportDate"></param>
        /// <param name="groupProductCode"></param>
        /// <param name="bondNumber"></param>
        /// <param name="refNumber"></param>
        /// <returns></returns>
        List<D54ViewModel> getD54(DateTime reportDate, string groupProductCode, string bondNumber, string refNumber);
    }
}