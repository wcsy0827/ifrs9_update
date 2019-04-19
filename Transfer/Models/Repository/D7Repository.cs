using Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Entity.Infrastructure;
using System.IO;
using System.Linq;
using System.Text;
using Transfer.Controllers;
using Transfer.Models.Interface;
using Transfer.Utility;
using Transfer.ViewModels;
using static Transfer.Enum.Ref;

namespace Transfer.Models.Repository
{
    public class D7Repository : ID7Repository
    {
        #region 其他
        protected Common common
        {
            get;
            private set;
        }

        protected Common.User _UserInfo
        {
            get;
            private set;
        }
        public D7Repository()
        {
            this.common = new Common();
            this._UserInfo = new Common.User();
        }
        #endregion

        #region Get D70
        public Tuple<bool, List<D70ViewModel>> getD70All(string type)
        {
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                var users = db.IFRS9_User.AsNoTracking().ToList();
                if (type == "D70")
                {
                    if (db.Watching_List_Parm.Any())
                    {
                      
                        return new Tuple<bool, List<D70ViewModel>>
                        (
                            true,
                            (
                                from q in db.Watching_List_Parm.AsNoTracking()
                                .AsEnumerable().OrderBy(x => x.Rule_ID)
                                select DbToD70ViewModel(q, users)
                            ).ToList()
                        );
                    }
                }
                else if (type == "D71")
                {
                    if (db.Warning_List_Parm.Any())
                    {
                        return new Tuple<bool, List<D70ViewModel>>
                        (
                            true,
                            (
                                from q in db.Warning_List_Parm.AsNoTracking()
                                .AsEnumerable().OrderBy(x => x.Rule_ID)
                                select DbToD71ViewModel(q, users)
                            ).ToList()
                        );
                    }
                }
            }

            return new Tuple<bool, List<D70ViewModel>>(true, new List<D70ViewModel>());
        }

        public Tuple<bool, List<D70ViewModel>> getD70(string type, D70ViewModel dataModel)
        {
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                var users = db.IFRS9_User.AsNoTracking().ToList();
                if (type == "D70")
                {
                    if (db.Watching_List_Parm.Any())
                    {
                        var query = from q in db.Watching_List_Parm.AsNoTracking()
                                    select q;

                        query = query.Where(x => x.Rule_ID.ToString() == dataModel.Rule_ID, dataModel.Rule_ID.IsNullOrWhiteSpace() == false)
                                     .Where(x => x.Status == dataModel.Status, dataModel.Status.IsNullOrWhiteSpace() == false)
                                     .Where(x => x.IsActive == dataModel.IsActive, dataModel.IsActive.IsNullOrWhiteSpace() == false);                     
                        return new Tuple<bool, List<D70ViewModel>>(query.Any(),
                            query.AsEnumerable().OrderBy(x => x.Rule_ID).Select(x => { return DbToD70ViewModel(x, users); }).ToList());
                    }
                }
                else if (type == "D71")
                {
                    if (db.Warning_List_Parm.Any())
                    {
                        var query = from q in db.Warning_List_Parm.AsNoTracking()
                                    select q;

                        query = query.Where(x => x.Rule_ID.ToString() == dataModel.Rule_ID, dataModel.Rule_ID.IsNullOrWhiteSpace() == false)
                                     .Where(x => x.Status == dataModel.Status, dataModel.Status.IsNullOrWhiteSpace() == false)
                                     .Where(x => x.IsActive == dataModel.IsActive, dataModel.IsActive.IsNullOrWhiteSpace() == false);
                        return new Tuple<bool, List<D70ViewModel>>(query.Any(),
                            query.AsEnumerable().OrderBy(x => x.Rule_ID).Select(x => { return DbToD71ViewModel(x, users); }).ToList());
                    }
                }
            }

            return new Tuple<bool, List<D70ViewModel>>(false, new List<D70ViewModel>());
        }
        #endregion

        #region Db 組成 D70ViewModel
        private D70ViewModel DbToD70ViewModel(Watching_List_Parm data,List<IFRS9_User> users)
        {
            D70ViewModel viewModel = new D70ViewModel();

            viewModel.Rule_ID = data.Rule_ID.ToString();
            viewModel.Basic_Pass_Check_Point = data.Basic_Pass_Check_Point;
            viewModel.Basic_Pass = data.Basic_Pass;
            viewModel.Rating_Check_Point = data.Rating_Check_Point;
            viewModel.Rating_Threshold = data.Rating_Threshold;
            viewModel.Data_Year = data.Data_Year;
            viewModel.Rating_Threshold_Map_Grade = data.Rating_Threshold_Map_Grade.ToString();
            viewModel.Rating_Threshold_Map_Grade_Adjust = data.Rating_Threshold_Map_Grade_Adjust.ToString();
            viewModel.Including_Ind_0 = data.Including_Ind_0;
            viewModel.Apply_Range_0 = data.Apply_Range_0;
            viewModel.Rating_from = data.Rating_from;
            viewModel.Rating_from_Map_Grade = data.Rating_from_Map_Grade.ToString();
            viewModel.Rating_from_Map_Grade_Adjust = data.Rating_from_Map_Grade_Adjust.ToString();
            viewModel.Rating_To = data.Rating_To;
            viewModel.Rating_To_Map_Grade = data.Rating_To_Map_Grade.ToString();
            viewModel.Rating_To_Map_Grade_Adjust = data.Rating_To_Map_Grade_Adjust.ToString();
            viewModel.Value_Change_Months = data.Value_Change_Months.ToString();
            viewModel.Including_Ind_1 = data.Including_Ind_1;
            viewModel.Apply_Range_1 = data.Apply_Range_1;
            viewModel.Spread_Change_Months = data.Spread_Change_Months.ToString();
            viewModel.Including_Ind_2 = data.Including_Ind_2;
            viewModel.Apply_Range_2 = data.Apply_Range_2;
            viewModel.Rule_setter = data.Rule_setter;
            viewModel.Rule_setting_Date = TypeTransfer.dateTimeNToString(data.Rule_setting_Date);
            viewModel.Auditor = data.Auditor;
            viewModel.Audit_Date = TypeTransfer.dateTimeNToString(data.Audit_Date);
            viewModel.Status = data.Status;
            viewModel.Rule_desc = data.Rule_desc;
            viewModel.IsActive = data.IsActive;

            return getNewD70ViewModel("D70", viewModel, users);
        }
        #endregion

        #region Db 組成 D71ViewModel
        private D70ViewModel DbToD71ViewModel(Warning_List_Parm data,List<IFRS9_User> users)
        {
            D70ViewModel viewModel = new D70ViewModel();

            viewModel.Rule_ID = data.Rule_ID.ToString();
            viewModel.Basic_Pass_Check_Point = data.Basic_Pass_Check_Point;
            viewModel.Basic_Pass = data.Basic_Pass;
            viewModel.Rating_Check_Point = data.Rating_Check_Point;
            viewModel.Rating_Threshold = data.Rating_Threshold;
            viewModel.Data_Year = data.Data_Year;
            viewModel.Rating_Threshold_Map_Grade = data.Rating_Threshold_Map_Grade.ToString();
            viewModel.Rating_Threshold_Map_Grade_Adjust = data.Rating_Threshold_Map_Grade_Adjust.ToString();
            viewModel.Including_Ind_0 = data.Including_Ind_0;
            viewModel.Apply_Range_0 = data.Apply_Range_0;
            viewModel.Rating_from = data.Rating_from;
            viewModel.Rating_from_Map_Grade = data.Rating_from_Map_Grade.ToString();
            viewModel.Rating_from_Map_Grade_Adjust = data.Rating_from_Map_Grade_Adjust.ToString();
            viewModel.Rating_To = data.Rating_To;
            viewModel.Rating_To_Map_Grade = data.Rating_To_Map_Grade.ToString();
            viewModel.Rating_To_Map_Grade_Adjust = data.Rating_To_Map_Grade_Adjust.ToString();
            viewModel.Value_Change_Months = data.Value_Change_Months.ToString();
            viewModel.Including_Ind_1 = data.Including_Ind_1;
            viewModel.Apply_Range_1 = data.Apply_Range_1;
            viewModel.Spread_Change_Months = data.Spread_Change_Months.ToString();
            viewModel.Including_Ind_2 = data.Including_Ind_2;
            viewModel.Apply_Range_2 = data.Apply_Range_2;
            viewModel.Rule_setter = data.Rule_setter;
            viewModel.Rule_setting_Date = TypeTransfer.dateTimeNToString(data.Rule_setting_Date);
            viewModel.Auditor = data.Auditor;
            viewModel.Audit_Date = TypeTransfer.dateTimeNToString(data.Audit_Date);
            viewModel.Status = data.Status;
            viewModel.Rule_desc = data.Rule_desc;
            viewModel.IsActive = data.IsActive;

            return getNewD70ViewModel("D71", viewModel, users);
        }
        #endregion

        private D70ViewModel getNewD70ViewModel(string tableId, D70ViewModel data,List<IFRS9_User> users)
        {
            if (data.Basic_Pass_Check_Point == "0")
            {
                data.Basic_Pass_Check_Point_Name = "0：本月報導日";
            }
            else if (data.Basic_Pass_Check_Point == "-1")
            {
                data.Basic_Pass_Check_Point_Name = "-1：上個月報導日";
            }
            else if (data.Basic_Pass_Check_Point == "-2")
            {
                data.Basic_Pass_Check_Point_Name = "-2：上上個月報導日";
            }

            if (data.Rating_Check_Point == "0")
            {
                data.Rating_Check_Point_Name = "0：本月報導日";
            }
            else if (data.Rating_Check_Point == "-1")
            {
                data.Rating_Check_Point_Name = "-1：上個月報導日";
            }
            else if (data.Rating_Check_Point == "-2")
            {
                data.Rating_Check_Point_Name = "-2：上上個月報導日";
            }

            if (data.Apply_Range_0 == "1")
            {
                data.Apply_Range_0_Name = "1：以上";
            }
            else if (data.Apply_Range_0 == "0")
            {
                data.Apply_Range_0_Name = "0：以下";
            }

            if (data.Apply_Range_1 == "1")
            {
                data.Apply_Range_1_Name = "1：以上";
            }
            else if (data.Apply_Range_1 == "0")
            {
                data.Apply_Range_1_Name = "0：以下";
            }

            if (data.Apply_Range_2 == "1")
            {
                data.Apply_Range_2_Name = "1：以上";
            }
            else if (data.Apply_Range_2 == "0")
            {
                data.Apply_Range_2_Name = "0：以下";
            }

            ProcessStatusList psl = new ProcessStatusList();
            List<SelectOption> selectOptionStatus = psl.statusOption;
            for (int i = 0; i < selectOptionStatus.Count; i++)
            {
                if (selectOptionStatus[i].Value == data.Status)
                {
                    data.Status_Name = selectOptionStatus[i].Text;
                    break;
                }
            }

            IsActiveList ial = new IsActiveList();
            selectOptionStatus = ial.isActiveOption;
            for (int i = 0; i < selectOptionStatus.Count; i++)
            {
                if (selectOptionStatus[i].Value == data.IsActive)
                {
                    data.IsActive_Name = selectOptionStatus[i].Text;
                    break;
                }
            }

            var UserData = users.Where(x => x.User_Account == data.Rule_setter).FirstOrDefault();
            if (UserData != null)
            {
                data.Rule_setter_Name = UserData.User_Name;
            }

            UserData = users.Where(x => x.User_Account == data.Auditor).FirstOrDefault();
            if (UserData != null)
            {
                data.Auditor_Name = UserData.User_Name;
            }         

            return data;
        }

        #region saveD70

        public MSGReturnModel saveD70(string type, string actionType, D70ViewModel dataModel)
        {
            MSGReturnModel result = new MSGReturnModel();

            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                try
                {
                    if (type == "D70")
                    {
                        Watching_List_Parm classData = new Watching_List_Parm();

                        if (actionType == "Add")
                        {
                            classData.IsActive = "N";
                        }
                        else if (actionType == "Modify")
                        {
                            classData = db.Watching_List_Parm.Where(x => x.Rule_ID.ToString() == dataModel.Rule_ID).FirstOrDefault();

                            classData.Auditor = "";
                            classData.Audit_Date = null;
                        }

                        classData.Basic_Pass_Check_Point = dataModel.Basic_Pass_Check_Point;
                        classData.Basic_Pass = dataModel.Basic_Pass;

                        classData.Rating_Check_Point = dataModel.Rating_Check_Point;

                        classData.Rating_Threshold = dataModel.Rating_Threshold;

                        classData.Data_Year = dataModel.Data_Year;

                        if (!dataModel.Rating_Threshold_Map_Grade.IsNullOrEmpty())
                        {
                            classData.Rating_Threshold_Map_Grade = int.Parse(dataModel.Rating_Threshold_Map_Grade);
                        }

                        if (!dataModel.Rating_Threshold_Map_Grade_Adjust.IsNullOrEmpty())
                        {
                            classData.Rating_Threshold_Map_Grade_Adjust = int.Parse(dataModel.Rating_Threshold_Map_Grade_Adjust);
                        }

                        classData.Including_Ind_0 = dataModel.Including_Ind_0;
                        classData.Apply_Range_0 = dataModel.Apply_Range_0;
                        classData.Rating_from = dataModel.Rating_from;

                        if (!dataModel.Rating_from_Map_Grade.IsNullOrEmpty())
                        {
                            classData.Rating_from_Map_Grade = int.Parse(dataModel.Rating_from_Map_Grade);
                        }

                        if (!dataModel.Rating_from_Map_Grade_Adjust.IsNullOrEmpty())
                        {
                            classData.Rating_from_Map_Grade_Adjust = int.Parse(dataModel.Rating_from_Map_Grade_Adjust);
                        }

                        classData.Rating_To = dataModel.Rating_To;

                        if (!dataModel.Rating_To_Map_Grade.IsNullOrEmpty())
                        {
                            classData.Rating_To_Map_Grade = int.Parse(dataModel.Rating_To_Map_Grade);
                        }

                        if (!dataModel.Rating_To_Map_Grade_Adjust.IsNullOrEmpty())
                        {
                            classData.Rating_To_Map_Grade_Adjust = int.Parse(dataModel.Rating_To_Map_Grade_Adjust);
                        }

                        if (!dataModel.Value_Change_Months.IsNullOrEmpty())
                        {
                            classData.Value_Change_Months = int.Parse(dataModel.Value_Change_Months);
                        }
                        classData.Including_Ind_1 = dataModel.Including_Ind_1;
                        classData.Apply_Range_1 = dataModel.Apply_Range_1;

                        if (!dataModel.Spread_Change_Months.IsNullOrEmpty())
                        {
                            classData.Spread_Change_Months = int.Parse(dataModel.Spread_Change_Months);
                        }
                        classData.Including_Ind_2 = dataModel.Including_Ind_2;
                        classData.Apply_Range_2 = dataModel.Apply_Range_2;

                        classData.Rule_setter = AccountController.CurrentUserInfo.Name;
                        classData.Rule_setting_Date = DateTime.Now.Date;

                        classData.Status = "0";

                        classData.Rule_desc = dataModel.Rule_desc;

                        if (actionType == "Add")
                        {
                            db.Watching_List_Parm.Add(classData);                   
                            db.SaveChanges();
                            result.RETURN_FLAG = true;
                            result.DESCRIPTION = Message_Type.insert_Success_Wait_Audit.GetDescription(type);
                        }
                    }
                    else if (type == "D71")
                    {
                        Warning_List_Parm classData = new Warning_List_Parm();

                        if (actionType == "Add")
                        {
                            classData.IsActive = "N";
                        }
                        else if (actionType == "Modify")
                        {
                            classData = db.Warning_List_Parm.Where(x => x.Rule_ID.ToString() == dataModel.Rule_ID).FirstOrDefault();

                            classData.Auditor = "";
                            classData.Audit_Date = null;
                        }

                        classData.Basic_Pass_Check_Point = dataModel.Basic_Pass_Check_Point;
                        classData.Basic_Pass = dataModel.Basic_Pass;

                        classData.Rating_Check_Point = dataModel.Rating_Check_Point;

                        classData.Rating_Threshold = dataModel.Rating_Threshold;

                        classData.Data_Year = dataModel.Data_Year;

                        if (!dataModel.Rating_Threshold_Map_Grade.IsNullOrEmpty())
                        {
                            classData.Rating_Threshold_Map_Grade = int.Parse(dataModel.Rating_Threshold_Map_Grade);
                        }

                        if (!dataModel.Rating_Threshold_Map_Grade_Adjust.IsNullOrEmpty())
                        {
                            classData.Rating_Threshold_Map_Grade_Adjust = int.Parse(dataModel.Rating_Threshold_Map_Grade_Adjust);
                        }

                        classData.Including_Ind_0 = dataModel.Including_Ind_0;
                        classData.Apply_Range_0 = dataModel.Apply_Range_0;
                        classData.Rating_from = dataModel.Rating_from;

                        if (!dataModel.Rating_from_Map_Grade.IsNullOrEmpty())
                        {
                            classData.Rating_from_Map_Grade = int.Parse(dataModel.Rating_from_Map_Grade);
                        }

                        if (!dataModel.Rating_from_Map_Grade_Adjust.IsNullOrEmpty())
                        {
                            classData.Rating_from_Map_Grade_Adjust = int.Parse(dataModel.Rating_from_Map_Grade_Adjust);
                        }

                        classData.Rating_To = dataModel.Rating_To;

                        if (!dataModel.Rating_To_Map_Grade.IsNullOrEmpty())
                        {
                            classData.Rating_To_Map_Grade = int.Parse(dataModel.Rating_To_Map_Grade);
                        }

                        if (!dataModel.Rating_To_Map_Grade_Adjust.IsNullOrEmpty())
                        {
                            classData.Rating_To_Map_Grade_Adjust = int.Parse(dataModel.Rating_To_Map_Grade_Adjust);
                        }

                        if (!dataModel.Value_Change_Months.IsNullOrEmpty())
                        {
                            classData.Value_Change_Months = int.Parse(dataModel.Value_Change_Months);
                        }
                        classData.Including_Ind_1 = dataModel.Including_Ind_1;
                        classData.Apply_Range_1 = dataModel.Apply_Range_1;

                        if (!dataModel.Spread_Change_Months.IsNullOrEmpty())
                        {
                            classData.Spread_Change_Months = int.Parse(dataModel.Spread_Change_Months);
                        }
                        classData.Including_Ind_2 = dataModel.Including_Ind_2;
                        classData.Apply_Range_2 = dataModel.Apply_Range_2;

                        classData.Rule_setter = AccountController.CurrentUserInfo.Name;
                        classData.Rule_setting_Date = DateTime.Now.Date;

                        classData.Status = "0";

                        classData.Rule_desc = dataModel.Rule_desc;

                        if (actionType == "Add")
                        {
                            db.Warning_List_Parm.Add(classData);
                            db.SaveChanges();
                            result.RETURN_FLAG = true;
                            result.DESCRIPTION = Message_Type.insert_Success_Wait_Audit.GetDescription(type);
                        }
                    }
                }
                catch (Exception ex)
                {
                    result.RETURN_FLAG = false;
                    result.DESCRIPTION = Message_Type.save_Fail.GetDescription(type, ex.Message);
                }
            }

            return result;
        }

        #endregion

        #region deleteD70
        public MSGReturnModel deleteD70(string type, string ruleID)
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                try
                {
                    if (type == "D70")
                    {
                        Watching_List_Parm oldData = db.Watching_List_Parm.Where(x => x.Rule_ID.ToString() == ruleID)
                                                                          .FirstOrDefault();
                        if (oldData.IsActive == "Y")
                        {
                            oldData.Rule_setter = AccountController.CurrentUserInfo.Name;
                            oldData.Rule_setting_Date = DateTime.Now.Date;
                            oldData.Auditor = "";
                            oldData.Audit_Date = null;
                            oldData.Status = "0";                           
                            db.SaveChanges();
                            result.RETURN_FLAG = true;
                            result.DESCRIPTION = Message_Type.delete_Success_Wait_Audit.GetDescription();
                        }
                        else
                        {
                            result.DESCRIPTION = Message_Type.already_IsActiveN.GetDescription();
                        }
                    }
                    else if (type == "D71")
                    {
                        Warning_List_Parm oldData = db.Warning_List_Parm.Where(x => x.Rule_ID.ToString() == ruleID)
                                                                        .FirstOrDefault();
                        if (oldData.IsActive == "Y")
                        {
                            oldData.Rule_setter = AccountController.CurrentUserInfo.Name;
                            oldData.Rule_setting_Date = DateTime.Now.Date;
                            oldData.Auditor = "";
                            oldData.Audit_Date = null;
                            oldData.Status = "0";
                            db.SaveChanges();
                            result.RETURN_FLAG = true;
                            result.DESCRIPTION = Message_Type.delete_Success_Wait_Audit.GetDescription();
                        }
                        else
                        {
                            result.DESCRIPTION = Message_Type.already_IsActiveN.GetDescription();
                        }
                    }
                }
                catch (DbUpdateException ex)
                {
                    result.RETURN_FLAG = false;
                    result.DESCRIPTION = Message_Type.save_Fail.GetDescription(null, ex.Message);
                }
            }
            return result;
        }
        #endregion

        #region sendD70ToAudit
        public MSGReturnModel sendD70ToAudit(string type, List<D70ViewModel> dataModelList)
        {
            MSGReturnModel result = new MSGReturnModel();
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                try
                {
                    if (type == "D70")
                    {
                        for (int i = 0; i < dataModelList.Count; i++)
                        {
                            int ruleID = int.Parse(dataModelList[i].Rule_ID);
                            Watching_List_Parm oldData = db.Watching_List_Parm.Where(x => x.Rule_ID == ruleID).FirstOrDefault();
                            oldData.Auditor = dataModelList[i].Auditor;
                            oldData.Status = "1";
                        }
                    }
                    else if (type == "D71")
                    {
                        for (int i = 0; i < dataModelList.Count; i++)
                        {
                            int ruleID = int.Parse(dataModelList[i].Rule_ID);
                            Warning_List_Parm oldData = db.Warning_List_Parm.Where(x => x.Rule_ID == ruleID).FirstOrDefault();
                            oldData.Auditor = dataModelList[i].Auditor;
                            oldData.Status = "1";
                        }
                    }

                    db.SaveChanges();
                    result.RETURN_FLAG = true;
                    result.DESCRIPTION = Message_Type.send_To_Audit_Success.GetDescription(type);
                }
                catch (DbUpdateException ex)
                {
                    result.RETURN_FLAG = false;
                    result.DESCRIPTION = Message_Type.send_To_Audit_Fail.GetDescription(type, ex.Message);
                }
            }

            return result;
        }
        #endregion

        #region D70Audit
        public MSGReturnModel D70Audit(string type, List<D70ViewModel> dataModelList)
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                try
                {
                    int ruleID = 0;

                    for (int i = 0; i < dataModelList.Count; i++)
                    {
                        ruleID = int.Parse(dataModelList[i].Rule_ID);

                        if (type == "D70")
                        {
                            Watching_List_Parm oldData = db.Watching_List_Parm.Where(x => x.Rule_ID == ruleID).FirstOrDefault();
                            oldData.Auditor = AccountController.CurrentUserInfo.Name;
                            oldData.Audit_Date = DateTime.Now.Date;

                            if (oldData.Status == "1")
                            {
                                //status == "2" => 2:複核完成
                                if (dataModelList[i].Status == "2" && oldData.IsActive == "N")//原為失效 => 新增案例
                                    oldData.IsActive = "Y";
                                else if (dataModelList[i].Status == "2" && oldData.IsActive == "Y")//原為有效 => 刪除(失效)案例
                                    oldData.IsActive = "N";
                                //status == "3" => 3:複核退回 資料保持不變
                            }
                            oldData.Status = dataModelList[i].Status;
                        }
                        else if (type == "D71")
                        {
                            Warning_List_Parm oldData = db.Warning_List_Parm.Where(x => x.Rule_ID == ruleID).FirstOrDefault();
                            oldData.Auditor = AccountController.CurrentUserInfo.Name;
                            oldData.Audit_Date = DateTime.Now.Date;
                            if (oldData.Status == "1")
                            {
                                //status == "2" => 2:複核完成
                                if (dataModelList[i].Status == "2" && oldData.IsActive == "N")//原為失效 => 新增案例
                                    oldData.IsActive = "Y";
                                else if (dataModelList[i].Status == "2" && oldData.IsActive == "Y")//原為有效 => 刪除(失效)案例
                                    oldData.IsActive = "N";
                                //status == "3" => 3:複核退回 資料保持不變
                            }
                            oldData.Status = dataModelList[i].Status;
                        }
                    }
                    db.SaveChanges();
                    result.RETURN_FLAG = true;
                    result.DESCRIPTION = Message_Type.Audit_Success.GetDescription(type);
                }
                catch (DbUpdateException ex)
                {
                    result.RETURN_FLAG = false;
                    result.DESCRIPTION = Message_Type.Audit_Fail.GetDescription(type, ex.Message);
                }
            }

            return result;
        }
        #endregion

        #region Get Data
        /// <summary>
        /// Get D72
        /// </summary>
        /// <returns></returns>
        public Tuple<string, List<D72ViewModel>> GetD72()
        {
            string resultMessage = string.Empty;
            List<D72ViewModel> data = new List<D72ViewModel>();
            try
            {
                using (IFRS9DBEntities db = new IFRS9DBEntities())
                {
                    var dbe = db.SMF_Group.AsNoTracking()
                        .OrderBy(x=>x.PRODUCT).AsEnumerable();
                    data = common.getViewModel(dbe, Table_Type.D72).OfType<D72ViewModel>().ToList();
                }
                if (!data.Any())
                    resultMessage = Message_Type.not_Find_Any.GetDescription(Table_Type.D72.tableNameGetDescription());
            }
            catch (Exception ex)
            {
                resultMessage = ex.exceptionMessage();
            }
            return new Tuple<string, List<D72ViewModel>>(resultMessage, data);
        }

        /// <summary>
        /// Get D73 Version
        /// </summary>
        /// <param name="date"></param>
        /// <returns></returns>
        public List<string> GetD73Veriosn(DateTime? date)
        {
            List<string> result = new List<string>();
            if (!date.HasValue)
                return result;
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                result = db.Scheduling_Report.AsNoTracking()
                    .Where(x => x.Report_Date == date.Value)
                    .Where(x => x.Version != null)
                    .Select(x => x.Version).Distinct()
                    .OrderByDescending(x=>x)
                    .AsEnumerable().Select(x=>x.Value.ToString())
                    .ToList();
            }
            return result;
        }

        /// <summary>
        /// 檢視 D73 檔案內容
        /// </summary>
        /// <param name="Id"></param>
        /// <returns></returns>
        public string GetD73FileLog(int Id)
        {
            string str = string.Empty;
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                var _d73 = db.Scheduling_Report.AsNoTracking()
                    .FirstOrDefault(x => x.ID == Id);
                if (_d73 != null)
                {
                    try
                    {
                        if (_d73.Delete_YN != "Y" && File.Exists(_d73.File_path))
                            using (StreamReader sr = new StreamReader(_d73.File_path))
                            {
                                str = sr.ReadToEnd();
                            }
                        else if(_d73.Delete_YN == "Y" && _d73.Delete_Date.HasValue)
                            str = $@"已於:{_d73.Delete_Date.Value.ToString("yyyy/MM/dd")} 刪除評估結果檔案,刪除帳號: {_d73.Processing_User}";
                    }
                    catch{}
                }
            }
            return str;
        }

        /// <summary>
        /// get D73 (排程查核結果紀錄檔)
        /// </summary>
        /// <param name="date">基準日</param>
        /// <param name="ver">版本</param>
        /// <returns></returns>
        public List<D73ViewModel> GetD73(DateTime? date,int? ver)
        {
            List<D73ViewModel> result = new List<D73ViewModel>();
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                result = db.Scheduling_Report.AsNoTracking()
                    .Where(x => x.Report_Date == date.Value, date.HasValue)
                    .Where(x => x.Version == ver.Value, ver.HasValue)
                    .OrderByDescending(x=>x.Create_Date)
                    .ThenByDescending(x=>x.File_Name)
                    .AsEnumerable()
                    .Select(x =>
                    new D73ViewModel
                    {
                        ID = x.ID.ToString(),
                        Report_Date = TypeTransfer.dateTimeNToString(x.Report_Date),
                        Version = TypeTransfer.intNToString(x.Version),
                        File_Name = x.File_Name,
                        File_path = x.File_path,
                        Create_Date = TypeTransfer.dateTimeToString(x.Create_Date,false),
                        Delete_YN = x.Delete_YN,
                        Delete_Date = TypeTransfer.dateTimeNToString(x.Delete_Date),
                        Processing_User = x.Processing_User
                    }).ToList();
            }
            return result;
        }

        /// <summary>
        /// 抓取啟用的主標尺對應檔年度
        /// </summary>
        /// <returns></returns>
        public string GetA51Year()
        {
            string result = string.Empty;

            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                result = db.Grade_Moody_Info.AsNoTracking().Where(x => x.Status == "1")
                    .FirstOrDefault()?.Data_Year;
            }
            return result;
        }

        /// <summary>
        /// 找尋設定的 通知名稱
        /// </summary>
        /// <param name="type"></param>
        /// <returns></returns>
        public List<SelectOption> GetNoticeName(string type)
        {
            
            List<SelectOption> result = new List<SelectOption>() { new SelectOption() { Text = " ", Value = " " } };
            if (type == "D74")
                result = new List<SelectOption>() { new SelectOption() { Text = "All", Value = "All" } };
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                result.AddRange(db.Notice_Info.AsNoTracking()
                    .AsEnumerable()
                    .Select(x => new SelectOption()
                    {
                        Text = $@"{x.Notice_Name}",
                        Value = x.Notice_ID.ToString()
                    }));
            }

            return result;
        }

        /// <summary>
        /// Get D74
        /// </summary>
        /// <param name="noticeName"></param>
        /// <returns></returns>
        public List<D74ViewModel> GetD74(string noticeID)
        {
            List<D74ViewModel> result = new List<D74ViewModel>();
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                result = db.Notice_Info.AsNoTracking()
                    .Where(x => x.Notice_ID.ToString() == noticeID, noticeID != "All")
                    .AsEnumerable()
                    .Select(x =>
                        new D74ViewModel()
                        {
                            Notice_ID = x.Notice_ID.ToString(),
                            Notice_Name = x.Notice_Name,
                            Notice_Memo = x.Notice_Memo,
                            Mail_Title = x.Mail_Title,
                            Mail_Msg = x.Mail_Msg,
                            IsActive = x.IsActive,
                            mailNum = x.Mail_Info?.Count.ToString(),
                            Create_User = x.Create_User,
                            Create_Date_Time = TypeTransfer.dateTimeNTimeSpanNToString(x.Create_Date, x.Create_Time),
                            LastUpdate_User = x.LastUpdate_User,
                            LastUpdate_Date_Time = TypeTransfer.dateTimeNTimeSpanNToString(x.LastUpdate_Date, x.LastUpdate_Time)
                        }
                    ).ToList();
            }
            return result;
        }

        /// <summary>
        ///  Get D74_1 
        /// </summary>
        /// <param name="noticeID"></param>
        /// <returns></returns>
        public List<D74_1ViewModel> GetD74_1(string noticeID)
        {
            List<D74_1ViewModel> result = new List<D74_1ViewModel>();
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                var notice = db.Notice_Info.AsNoTracking()
                    .FirstOrDefault(x => x.Notice_ID.ToString() == noticeID);
                if (notice != null)
                {
                    result.AddRange(notice.Mail_Info.Select(x => new D74_1ViewModel()
                    {
                        ID = x.ID.ToString(),
                        Notice_ID = x.Notice_ID.ToString(),
                        Recipient_Department = x.Recipient_Department,
                        Recipient_mail = x.Recipient_mail,
                        Recipient_Name = x.Recipient_Name,
                        IsActive = x.IsActive,
                        Create_User = x.Create_User,
                        Create_Date_Time = TypeTransfer.dateTimeNTimeSpanNToString(x.Create_Date, x.Create_Time),
                        LastUpdate_User = x.LastUpdate_User,
                        LastUpdate_Date_Time = TypeTransfer.dateTimeNTimeSpanNToString(x.LastUpdate_Date, x.LastUpdate_Time)
                    }));
                }
            }
            return result;
        }

        #endregion

        #region Save Db
        /// <summary>
        /// Save D72
        /// </summary>
        /// <param name="datas"></param>
        /// <returns></returns>
        public MSGReturnModel SaveD72(List<D72ViewModel> datas)
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            if (!datas.Any())
            {
                result.DESCRIPTION = Message_Type.not_Find_Any.GetDescription();
                return result;
            }
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                StringBuilder sb = new StringBuilder();
                db.SMF_Info.RemoveRange(db.SMF_Info);
                sb.Append($@"
                delete SMF_Group ;  ");
                datas.ForEach(x =>
                {
                    sb.Append($@"
INSERT INTO [SMF_Group]
           ([PRODUCT]
           ,[Product_Group]
           ,[Product_Group_kind]
           ,[Memo])
     VALUES
           ({x.PRODUCT.stringToStrSql()}
           ,{x.Product_Group.stringToStrSql()}
           ,{x.Product_Group_kind.stringToStrSql()}
           ,{x.Memo.stringToStrSql()} ) ; ");
                });
                try
                {
                    db.Database.ExecuteSqlCommand(sb.ToString());
                    result.RETURN_FLAG = true;
                    result.DESCRIPTION = Message_Type.save_Success.GetDescription();
                }
                catch (DbUpdateException ex)
                {
                    result.DESCRIPTION = ex.exceptionMessage();
                }
            }
            return result;
        }

        /// <summary>
        /// 刪除 D73 Log檔案
        /// </summary>
        /// <param name="ID"></param>
        /// <returns></returns>
        public MSGReturnModel DelD73(string ID)
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                var tableName = Table_Type.D73.tableNameGetDescription();
                var Id = 0;
                if (!Int32.TryParse(ID, out Id))
                {
                    result.DESCRIPTION = Message_Type.parameter_Error.GetDescription(tableName);
                    return result;
                }
                var _D73 = db.Scheduling_Report.FirstOrDefault(x => x.ID == Id && x.Delete_YN != "Y");
                if (_D73 == null || !File.Exists(_D73.File_path))
                {
                    result.DESCRIPTION = Message_Type.already_Change.GetDescription(tableName);
                    return result;
                }

                _D73.Delete_YN = "Y";
                _D73.Delete_Date = DateTime.Now.Date;
                _D73.Processing_User = _UserInfo._user;
                try
                {
                    System.IO.File.Delete(_D73.File_path);
                    db.SaveChanges();
                    result.RETURN_FLAG = true;
                    result.DESCRIPTION = Message_Type.delete_Success.GetDescription(tableName);
                }
                catch (Exception ex)
                {
                    result.DESCRIPTION = ex.exceptionMessage();
                }
            }
            return result;
        }

        /// <summary>
        /// 新增通知設定檔
        /// </summary>
        /// <param name="data"></param>
        /// <returns></returns>
        public MSGReturnModel AddD74(D74ViewModel data)
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            result.DESCRIPTION = Message_Type.save_Fail.GetDescription();
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                var _Notice_Name = data.Notice_Name?.Trim();
                if (db.Notice_Info.AsNoTracking().Any(x => x.Notice_Name == _Notice_Name))
                {
                    result.DESCRIPTION = $"已設定過相同的名稱:{_Notice_Name}";
                    return result;
                }

                db.Notice_Info.Add(
                    new Notice_Info
                    {
                        Notice_Name = data.Notice_Name?.Trim(),
                        Notice_Memo = data.Notice_Memo?.Trim(),
                        Mail_Title = data.Mail_Title?.Trim(),
                        Mail_Msg = data.Mail_Msg?.Trim(),
                        IsActive = data.IsActive,
                        Create_User = _UserInfo._user,
                        Create_Date  = _UserInfo._date,
                        Create_Time = _UserInfo._time,
                        LastUpdate_User = _UserInfo._user,
                        LastUpdate_Date = _UserInfo._date,
                        LastUpdate_Time = _UserInfo._time
                    });
                try
                {
                    db.SaveChanges();
                    result.RETURN_FLAG = true;
                    result.DESCRIPTION = Message_Type.save_Success.GetDescription();
                }
                catch (DbUpdateException ex)
                {
                    result.DESCRIPTION = Message_Type.save_Fail.GetDescription(null,ex.exceptionMessage());
                }
                catch (Exception ex) {
                    result.DESCRIPTION = Message_Type.save_Fail.GetDescription(null, ex.exceptionMessage());
                }
            }
            return result;
        }

        /// <summary>
        /// 修改通知設定檔
        /// </summary>
        /// <param name="data"></param>
        /// <returns></returns>
        public MSGReturnModel SaveD74(D74ViewModel data)
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            result.DESCRIPTION = Message_Type.not_Find_Any.GetDescription();
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                var _data = db.Notice_Info
                    .FirstOrDefault(x => x.Notice_ID.ToString() == data.Notice_ID);
                if (_data != null)
                {
                    _data.Notice_Name = data.Notice_Name?.Trim();
                    _data.Notice_Memo = data.Notice_Memo?.Trim();
                    _data.Mail_Title = data.Mail_Title?.Trim();
                    _data.Mail_Msg = data.Mail_Msg?.Trim();
                    _data.IsActive = data.IsActive;
                    _data.LastUpdate_User = _UserInfo._user;
                    _data.LastUpdate_Date = _UserInfo._date;
                    _data.LastUpdate_Time = _UserInfo._time;
                }
                else
                {
                    result.DESCRIPTION = Message_Type.not_Find_Any.GetDescription();
                    return result;
                }
                try
                {
                    db.SaveChanges();
                    result.RETURN_FLAG = true;
                    result.DESCRIPTION = Message_Type.update_Success.GetDescription();
                }
                catch (DbUpdateException ex)
                {
                    result.DESCRIPTION = Message_Type.update_Fail.GetDescription(null, ex.exceptionMessage());
                }
                catch (Exception ex)
                {
                    result.DESCRIPTION = Message_Type.update_Fail.GetDescription(null, ex.exceptionMessage());
                }
            }
            return result;
        }

        /// <summary>
        /// 新增郵件設定檔
        /// </summary>
        /// <param name="data"></param>
        /// <returns></returns>
        public MSGReturnModel AddD74_1(D74_1ViewModel data)
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                var addNotice_ID = TypeTransfer.stringToInt(data.Notice_ID);
                var addMail = data.Recipient_mail?.Trim();
                if (db.Mail_Info.AsNoTracking().Any(x => x.Notice_ID == addNotice_ID &&
                 x.Recipient_mail == addMail))
                {
                    result.DESCRIPTION = $"已設定過相同的收件者Mail:{addMail}";
                    return result;
                }

                db.Mail_Info.Add(new Mail_Info()
                {
                    Notice_ID = addNotice_ID,
                    Recipient_Department = data.Recipient_Department?.Trim(),
                    Recipient_Name = data.Recipient_Name?.Trim(),
                    Recipient_mail = data.Recipient_mail?.Trim(),
                    IsActive = data.IsActive,
                    Create_User = _UserInfo._user,
                    Create_Date = _UserInfo._date,
                    Create_Time = _UserInfo._time,
                    LastUpdate_User = _UserInfo._user,
                    LastUpdate_Date = _UserInfo._date,
                    LastUpdate_Time = _UserInfo._time
                });
                try
                {
                    db.SaveChanges();
                    result.RETURN_FLAG = true;
                    result.DESCRIPTION = Message_Type.save_Success.GetDescription();
                }
                catch (DbUpdateException ex)
                {
                    result.DESCRIPTION = Message_Type.save_Fail.GetDescription(null, ex.exceptionMessage());
                }
                catch (Exception ex)
                {
                    result.DESCRIPTION = Message_Type.save_Fail.GetDescription(null, ex.exceptionMessage());
                }
            }

            return result;
        }

        /// <summary>
        /// 修改郵件設定檔
        /// </summary>
        /// <param name="data"></param>
        /// <returns></returns>
        public MSGReturnModel SaveD74_1( D74_1ViewModel data)
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            result.DESCRIPTION = Message_Type.not_Find_Any.GetDescription();
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                var updateID = TypeTransfer.stringToInt(data.ID);
                var updateMail = data.Recipient_mail?.Trim();
                var _update = db.Mail_Info.FirstOrDefault(x => x.ID == updateID);
                if (_update != null)
                {
                    _update.Recipient_mail = data.Recipient_mail?.Trim();
                    _update.Recipient_Department = data.Recipient_Department?.Trim();
                    _update.Recipient_Name = data.Recipient_Name?.Trim();
                    _update.IsActive = data.IsActive;
                    _update.LastUpdate_User = _UserInfo._user;
                    _update.LastUpdate_Date = _UserInfo._date;
                    _update.LastUpdate_Time = _UserInfo._time;

                    try {
                        db.SaveChanges();
                        result.RETURN_FLAG = true;
                        result.DESCRIPTION = Message_Type.update_Success.GetDescription();
                    }
                    catch (DbUpdateException ex)
                    {
                        result.DESCRIPTION = Message_Type.update_Fail.GetDescription(null, ex.exceptionMessage());
                    }
                    catch (Exception ex)
                    {
                        result.DESCRIPTION = Message_Type.update_Fail.GetDescription(null, ex.exceptionMessage());
                    }
                }
            }

            return result;
        }

        /// <summary>
        /// 刪除郵件設定檔
        /// </summary>
        /// <param name="data"></param>
        /// <returns></returns>
        public MSGReturnModel DeleD74_1(D74_1ViewModel data)
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            result.DESCRIPTION = Message_Type.not_Find_Any.GetDescription();
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                var updateNotice_ID = TypeTransfer.stringToInt(data.Notice_ID);
                var updateMail = data.Recipient_mail?.Trim();
                var _update = db.Mail_Info.FirstOrDefault(x => x.Notice_ID == updateNotice_ID &&
                x.Recipient_mail == updateMail);
                if (_update != null)
                {
                    db.Mail_Info.Remove(_update);
                    try
                    {
                        db.SaveChanges();
                        result.RETURN_FLAG = true;
                        result.DESCRIPTION = Message_Type.delete_Success.GetDescription();
                    }
                    catch (DbUpdateException ex)
                    {
                        result.DESCRIPTION = Message_Type.delete_Fail.GetDescription(null, ex.exceptionMessage());
                    }
                    catch (Exception ex)
                    {
                        result.DESCRIPTION = Message_Type.delete_Fail.GetDescription(null, ex.exceptionMessage());
                    }
                }
            }

            return result;
        }

        #endregion

        #region Excel 部分
        /// <summary>
        /// Excel資料 轉 D72ViewModel
        /// </summary>
        /// <param name="pathType"></param>
        /// <param name="stream"></param>
        /// <returns></returns>
        public Tuple<string, List<D72ViewModel>> GetD72Excel(string pathType, Stream stream)
        {
            List<D72ViewModel> dataModel = new List<D72ViewModel>();
            string message = string.Empty;
            DataSet resultData = new DataSet();
            try
            {
                IExcelDataReader reader = null;
                switch (pathType) //判斷型別
                {
                    case "xls":
                        reader = ExcelReaderFactory.CreateBinaryReader(stream);
                        break;

                    case "xlsx":
                        reader = ExcelReaderFactory.CreateOpenXmlReader(stream);
                        break;
                }
                reader.IsFirstRowAsColumnNames = true;
                resultData = reader.AsDataSet();
                reader.Close();
                if (resultData.Tables[0].Rows.Count > 0) //判斷有無資料
                {
                    dataModel = common.getViewModel(resultData.Tables[0], Table_Type.D72)
                                      .Cast<D72ViewModel>().ToList();
                }
            }
            catch (Exception ex)
            {
                message = ex.exceptionMessage();
            }
            if (!dataModel.Any())
                message = Message_Type.not_Find_Any.GetDescription();
            return new Tuple<string, List<D72ViewModel>>(message, dataModel);
        }
        #endregion
    }
}