using System;
using System.Collections.Generic;
using System.IO;
using Transfer.Utility;
using Transfer.ViewModels;

namespace Transfer.Models.Interface
{
    public interface IA6Repository
    {
        Tuple<bool, List<A62ViewModel>> GetA62(string dataYear);

        /// <summary>
        /// Get A62 Search Year
        /// </summary>
        /// <returns></returns>
        List<string> GetA62SearchYear(string Status = "All");

        /// <summary>
        /// Excel資料 轉 Exhibit29Model
        /// </summary>
        /// <param name="pathType"></param>
        /// <param name="stream"></param>
        /// <returns></returns>
        List<Exhibit7Model> getExcel(string pathType, Stream stream);

        MSGReturnModel DownLoadA62Excel(string path, List<A62ViewModel> dbDatas);

        /// <summary>
        /// save A62 To Db
        /// </summary>
        /// <param name="dataModel">Exhibit7Model</param>
        /// <returns></returns>
        MSGReturnModel saveA62(List<Exhibit7Model> dataModel);

        List<A63ViewModel> getA63Excel(string pathType, Stream stream);

        MSGReturnModel saveA63(List<A63ViewModel> dataModel);

        MSGReturnModel A62Audit(List<A62ViewModel> dataModel);
    }
}