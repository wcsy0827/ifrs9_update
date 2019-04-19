using System;
using System.Collections.Generic;
using System.IO;
using Transfer.Utility;
using Transfer.ViewModels;

namespace Transfer.Models.Interface
{
    public interface IA8Repository
    {
        /// <summary>
        /// Get A81 資料
        /// </summary>
        /// <returns></returns>
        Tuple<bool, List<A81ViewModel>> GetA81();

        /// <summary>
        /// Get A82 資料
        /// </summary>
        /// <returns></returns>
        Tuple<bool, List<Moody_Quartly_PD_Info>> GetA82();

        /// <summary>
        /// Get A83 資料
        /// </summary>
        /// <returns></returns>
        Tuple<bool, List<Moody_Predit_PD_Info>> GetA83();

        /// <summary>
        /// Excel資料 轉 Exhibit10Model
        /// </summary>
        /// <param name="pathType"></param>
        /// <param name="stream"></param>
        /// <returns></returns>
        List<Exhibit10Model> getExcel(string pathType, Stream stream);

        /// <summary>
        /// save A81.A82.A83 資料
        /// </summary>
        /// <param name="type"></param>
        /// <param name="dataModel"></param>
        /// <returns></returns>
        MSGReturnModel SaveA8(string type, List<Exhibit10Model> dataModel);
    }
}