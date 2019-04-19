using System;
using System.Collections.Generic;
using Transfer.Utility;
using Transfer.ViewModels;
using static Transfer.Enum.Ref;

namespace Transfer.Models.Interface
{
    public interface ISystemForITRepository
    {
        /// <summary>
        /// 讀取每一位使用者的 menu
        /// </summary>
        /// <param name="userName"></param>
        /// <returns></returns>
        List<CheckBoxListInfo> getMenu(string userName);

        /// <summary>
        /// get can modify menu 的 帳號
        /// </summary>
        /// <param name="type">menu=>設定畫面的帳號,log=>觀察log的帳號</param>
        /// <param name="searchAll">log時使用(查詢包含失效(刪除))</param>
        /// <param name="adminAccount">管理者帳號</param>
        /// <returns></returns>
        List<Tuple<string, string>> getUser(string type, bool searchAll, string adminAccount);

        /// <summary>
        /// get ProductCode
        /// </summary>
        /// <param name="adminAccount"></param>
        /// <returns></returns>
        List<Tuple<string, string>> getProductCode(string adminAccount);

        /// <summary>
        /// save user menu 設定
        /// </summary>
        /// <param name="menuSub"></param>
        /// <param name="userName"></param>
        /// <returns></returns>
        MSGReturnModel saveMenu(List<CheckBoxListInfo> menuSub, string userName);

        /// <summary>
        /// get 管理帳號
        /// </summary>
        /// <param name="account">查詢帳號(沒有為全查)</param>
        /// <param name="admin">管理者帳號</param>
        /// <returns></returns>
        List<AccountViewModel> getAccount(string account, string admin);

        /// <summary>
        /// 查詢帳號使用紀錄
        /// </summary>
        /// <param name="account">查詢的條件</param>
        /// <param name="type">查詢的log,User=IFRS9_User_Log,Browser=IFRS9_Browse_Log,Event=IFRS9_Event_Log</param>
        /// <param name="range">查詢的範圍 今天,7天,全部 (tpye = user 使用)</param>
        /// <returns></returns>
        List<AccountLogViewModel> getAccountLog(AccountLogViewModel model, string type, string range);

        /// <summary>
        /// 查詢信評設定資料
        /// </summary>
        /// <param name="ProductCode"></param>
        /// <param name="tableId"></param>
        /// <param name="type"></param>
        /// <param name="effective"></param>
        /// <returns></returns>
        List<SetAssessmentViewModel> getAssessment(string productCode, string tableId, SetAssessmentType type, bool effective = true);

        /// <summary>
        /// get assessment can add Auditor or Presented User
        /// </summary>
        /// <param name="productCode"></param>
        /// <param name="tableId"></param>
        /// <param name="type"></param>
        /// <returns></returns>
        List<Tuple<string, string>> getAssessmentAddUser(string productCode, string tableId, SetAssessmentType type);

        /// <summary>
        /// save 帳號
        /// </summary>
        /// <param name="action">處理動作</param>
        /// <param name="data">該筆資料</param>
        /// <returns></returns>
        MSGReturnModel saveAccount(string action, IFRS9_User data);

        /// <summary>
        /// 信評相關新增
        /// </summary>
        /// <param name="model"></param>
        /// <param name="type"></param>
        /// <returns></returns>
        MSGReturnModel AssessmentAdd(SetAssessmentViewModel model, string type);

        /// <summary>
        /// 信評相關刪除
        /// </summary>
        /// <param name="model"></param>
        /// <param name="type"></param>
        /// <returns></returns>
        MSGReturnModel AssessmentDel(List<SetAssessmentViewModel> model, string type);

        Tuple<bool, List<MenuMainViewModel>> getMenuMainAll();
        Tuple<bool, List<MenuMainViewModel>> getMenuMain(MenuMainViewModel dataModel);
        MSGReturnModel saveMenuMain(string actionType, MenuMainViewModel dataModel);
        MSGReturnModel deleteMenuMain(string menu);

        Tuple<bool, List<MenuSubViewModel>> getMenuSubAll();
        Tuple<bool, List<MenuSubViewModel>> getMenuSub(MenuSubViewModel dataModel);
        MSGReturnModel saveMenuSub(string actionType, MenuSubViewModel dataModel);
        MSGReturnModel deleteMenuSub(string menu);

        List<TaskScheduleViewModel> getTaskSchedule(DateTime? dt, bool addFlag = true);
        MSGReturnModel restartTaskSchedule(string triggerTask, string reportDate);
    }
}