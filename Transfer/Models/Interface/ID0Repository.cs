using System;
using System.Collections.Generic;
using Transfer.Utility;
using Transfer.ViewModels;

namespace Transfer.Models.Interface
{
    public interface ID0Repository
    {
        /// <summary>
        /// get Db Group_Product data by debtType
        /// </summary>
        /// <param name="debtType">1.房貸 4.債券</param>
        /// <returns></returns>
        Tuple<bool, List<GroupProductViewModel>> getGroupProductByDebtType(string debtType);

        Tuple<bool, List<D03ViewModel>> getD03All();
        Tuple<bool, List<D03ViewModel>> getD03(D03ViewModel dataModel);
        MSGReturnModel saveD03(string actionType, D03ViewModel dataModel);
        MSGReturnModel deleteD03(string parmID);
        MSGReturnModel sendD03ToAudit(string parmID, string auditor);
        MSGReturnModel D03Audit(string parmID, string status);

        /// <summary>
        /// get Db D05 all data
        /// </summary>
        /// <returns></returns>
        Tuple<bool, List<D05ViewModel>> getD05All(string debtType);

        /// <summary>
        /// get Db D05 data
        /// </summary>
        /// <returns></returns>
        Tuple<bool, List<D05ViewModel>> getD05(string debtType, string groupProductCode, string productCode, string processingDate1, string processingDate2);

        /// <summary>
        /// save D05 To Db
        /// </summary>
        /// <param name="debtType"></param>
        /// <param name="actionType"></param>
        /// <param name="dataModel">D05ViewModel</param>
        /// <returns></returns>
        MSGReturnModel saveD05(string debtType, string actionType, D05ViewModel dataModel);

        /// <summary>
        /// delete D05
        /// </summary>
        /// <param name="productCode">產品</param>
        /// <returns></returns>
        MSGReturnModel deleteD05(string groupProductCode,string productCode);

        /// <summary>
        /// get Db D01 all data
        /// </summary>
        /// <param name="debtType">1.房貸 4.債券</param>
        /// <returns></returns>
        Tuple<bool, List<D01ViewModel>> getD01All(string debtType);

        /// <summary>
        /// get Db D01 data
        /// </summary>
        /// <param name="dataModel"></param>
        /// <returns></returns>
        Tuple<bool, List<D01ViewModel>> getD01(D01ViewModel dataModel);

        /// <summary>
        /// save D01 To Db
        /// </summary>
        /// <param name="dataModel">D01ViewModel</param>
        /// <returns></returns>
        MSGReturnModel saveD01(D01ViewModel dataModel);

        /// <summary>
        /// delete D01
        /// </summary>
        /// <param name="prjid">專案名稱</param>
        /// <param name="flowid">流程名稱</param>
        /// <returns></returns>
        MSGReturnModel deleteD01(string prjid, string flowid);

        /// <summary>
        /// get Group_Product all data
        /// </summary>
        /// <returns></returns>
        Tuple<bool, List<GroupProductViewModel>> getGroupProductAll(string debtType);

        /// <summary>
        /// get Db Group_Product data
        /// </summary>
        /// <returns></returns>
        Tuple<bool, List<GroupProductViewModel>> getGroupProduct(string debtType, string groupProductCode, string groupProductName);

        /// <summary>
        /// save Group_Product To Db
        /// </summary>
        /// <param name="debtType"></param>
        /// <param name="groupProductCode"></param>
        /// <param name="groupProductName">GroupProductViewModel</param>
        /// <returns></returns>
        MSGReturnModel saveGroupProduct(string debtType, string actionType, GroupProductViewModel dataModel);

        /// <summary>
        /// delete Group_Product
        /// </summary>
        /// <param name="groupProductCode">產品群代碼</param>
        /// <returns></returns>
        MSGReturnModel deleteGroupProduct(string groupProductCode);

        /// <summary>
        /// get Db D02 data
        /// </summary>
        /// <param name="dataModel"></param>
        /// <returns></returns>
        Tuple<bool, List<D02ViewModel>> getD02(D02ViewModel dataModel);

        MSGReturnModel saveD02(string reportDate);

        void saveD02AfterKRisk(string reportDate);

    }
}