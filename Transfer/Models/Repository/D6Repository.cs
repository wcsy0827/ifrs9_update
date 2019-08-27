using System;
using System.Collections.Generic;
using Transfer.Models.Interface;
using Transfer.ViewModels;
using System.Linq;
using Transfer.Utility;
using System.Data.Entity.Infrastructure;
using static Transfer.Enum.Ref;
using System.Data;
using Transfer.Controllers;
using System.Text;
using System.IO;
using System.Data.SqlClient;

namespace Transfer.Models.Repository
{
    public class D6Repository : ID6Repository
    {
        private string QuantifyAssessmentKind =
            AssessmentKind_Type.Quantify.GetDescription(); //量化衡量 
        private string QuantifyAssessmentStage = "2"; //量化衡量 Assessment_Stage
        private string QualitativeAssessmentKind =
            AssessmentKind_Type.Qualitative.GetDescription(); //質化衡量

        private string checkItemCode = "Pass_Count_";

        #region 其他

        public D6Repository()
        {
            this.common = new Common();
            this._UserInfo = new Common.User();
        }

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
        #endregion 其他

        #region Get IFRS9_Assessment_Presented_Config
        public Tuple<bool, List<IFRS9_User>> getIFRS9_Assessment_Presented_Config(string groupProductCode, string tableId, string userAccount)
        {
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                var query = common.getAssessmentInfo(groupProductCode, tableId, SetAssessmentType.Presented);
                if (query.Any(x=>x.User_Account == userAccount))
                {
                    return new Tuple<bool, List<IFRS9_User>>
                    (
                        true,
                        query
                    );
                }
            }

            return new Tuple<bool, List<IFRS9_User>>(true, new List<IFRS9_User>());
        }
        #endregion

        #region Get IFRS9_Assessment_Auditor_Config
        public Tuple<bool, List<IFRS9_User>> getIFRS9_Assessment_Auditor_Config(string groupProductCode, string tableId, string userAccount)
        {
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                var query = common.getAssessmentInfo(groupProductCode, tableId, SetAssessmentType.Auditor);
                if (query.Any())
                {
                    return new Tuple<bool, List<IFRS9_User>>
                    (
                        true,
                        query
                    );
                }
            }

            return new Tuple<bool, List<IFRS9_User>>(true, new List<IFRS9_User>());
        }
        #endregion

        #region getD60All
        public Tuple<bool, List<D60ViewModel>> getD60All()
        {
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                var users = db.IFRS9_User.AsNoTracking().ToList();
                if (db.Bond_Rating_Parm.Any())
                {
                    return new Tuple<bool, List<D60ViewModel>>
                                (
                                    true,
                                    (
                                        from q in db.Bond_Rating_Parm.AsNoTracking()
                                                    .AsEnumerable()
                                        select DbToD60Model(q, users)
                                    ).ToList()
                                );
                }
            }

            return new Tuple<bool, List<D60ViewModel>>(true, new List<D60ViewModel>());
        }
        #endregion

        #region DbToD60Model
        private D60ViewModel DbToD60Model(Bond_Rating_Parm data,List<IFRS9_User> users)
        {
            string ruleSetterName = "";
            string auditorName = "";
            string statusName = "";
            string isActiveName = "";

            ProcessStatusList psl = new ProcessStatusList();
            List<SelectOption> selectOptionStatus = psl.statusOption;
            for (int i = 0; i < selectOptionStatus.Count; i++)
            {
                if (selectOptionStatus[i].Value == data.Status)
                {
                    statusName = selectOptionStatus[i].Text;
                    break;
                }
            }

            IsActiveList ial = new IsActiveList();
            selectOptionStatus = ial.isActiveOption;
            for (int i = 0; i < selectOptionStatus.Count; i++)
            {
                if (selectOptionStatus[i].Value == data.IsActive)
                {
                    isActiveName = selectOptionStatus[i].Text;
                    break;
                }
            }

            var UserData = users.Where(x => x.User_Account == data.Rule_setter).FirstOrDefault();
            if (UserData != null)
            {
                ruleSetterName = UserData.User_Name;
            }

            UserData = users.Where(x => x.User_Account == data.Auditor).FirstOrDefault();
            if (UserData != null)
            {
                auditorName = UserData.User_Name;
            }
            
            return new D60ViewModel()
            {
                Parm_ID = data.Parm_ID.ToString(),
                Rating_Priority = data.Rating_Priority.ToString(),
                Rating_Object = data.Rating_Object,
                Rating_Org_Area = data.Rating_Org_Area,
                Rating_Selection = data.Rating_Selection,
                Rule_setter = data.Rule_setter,
                Rule_setter_Name = ruleSetterName,
                Rule_setting_Date = data.Rule_setting_Date.ToString("yyyy/MM/dd"),
                Auditor = data.Auditor,
                Auditor_Name = auditorName,
                Audit_Date = TypeTransfer.dateTimeNToString(data.Audit_Date),
                Status = data.Status,
                Status_Name = statusName,
                IsActive = data.IsActive,
                IsActive_Name = isActiveName
            };
        }
        #endregion

        #region  getD60
        public Tuple<bool, List<D60ViewModel>> getD60(D60ViewModel dataModel)
        {
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                if (db.Bond_Rating_Parm.Any())
                {
                    var query = db.Bond_Rating_Parm.AsNoTracking()
                                  .Where(x => x.Parm_ID.ToString() == dataModel.Parm_ID, dataModel.Parm_ID.IsNullOrWhiteSpace() == false)
                                  .Where(x => x.Rating_Priority.ToString() == dataModel.Rating_Priority, dataModel.Rating_Priority.IsNullOrWhiteSpace() == false)
                                  .Where(x => x.Rating_Object.Contains(dataModel.Rating_Object), dataModel.Rating_Object.IsNullOrWhiteSpace() == false)
                                  .Where(x => x.Rating_Org_Area == dataModel.Rating_Org_Area, dataModel.Rating_Org_Area.IsNullOrWhiteSpace() == false)
                                  .Where(x => x.Rating_Selection == dataModel.Rating_Selection, dataModel.Rating_Selection.IsNullOrWhiteSpace() == false)
                                  .Where(x => x.Status == dataModel.Status, dataModel.Status.IsNullOrWhiteSpace() == false)
                                  .Where(x => x.IsActive == dataModel.IsActive, dataModel.IsActive.IsNullOrWhiteSpace() == false);

                    if (dataModel.Rule_setting_Date.IsNullOrWhiteSpace() == false)
                    {
                        DateTime dt = DateTime.Parse(dataModel.Rule_setting_Date);
                        query = query.Where(x => x.Rule_setting_Date == dt);
                    }
                    var users = db.IFRS9_User.AsNoTracking().ToList();
                    return new Tuple<bool, List<D60ViewModel>>(query.Any(), query
                                                                            .AsEnumerable()
                                                                            .Select(x => { return DbToD60Model(x, users); }).ToList());
                }
            }

            return new Tuple<bool, List<D60ViewModel>>(false, new List<D60ViewModel>());
        }
        #endregion

        #region saveD60
        public MSGReturnModel saveD60(string actionType, D60ViewModel dataModel)
        {
            MSGReturnModel result = new MSGReturnModel();

            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                try
                {
                    if (actionType == "Add")
                    {
                        Bond_Rating_Parm addData = new Bond_Rating_Parm();

                        addData.Rating_Priority = int.Parse(dataModel.Rating_Priority);
                        addData.Rating_Object = dataModel.Rating_Object;
                        addData.Rating_Org_Area = dataModel.Rating_Org_Area;
                        addData.Rating_Selection = dataModel.Rating_Selection;
                        addData.Rule_setter = AccountController.CurrentUserInfo.Name;
                        addData.Rule_setting_Date = DateTime.Now.Date;
                        addData.Auditor = "";
                        addData.Audit_Date = null;
                        addData.Status = "0";
                        addData.IsActive = "N";

                        db.Bond_Rating_Parm.Add(addData);
                    }
                    else if (actionType == "Modify")
                    {
                    }

                    db.SaveChanges(); //Save

                    result.RETURN_FLAG = true;
                    result.DESCRIPTION = Message_Type.insert_Success_Wait_Audit.GetDescription();
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

        #region deleteD60
        public MSGReturnModel deleteD60(string parmID)
        {
            MSGReturnModel result = new MSGReturnModel();

            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                try
                {
                    Bond_Rating_Parm dataEdit = db.Bond_Rating_Parm
                                                  .Where(x => x.Parm_ID.ToString() == parmID)
                                                  .FirstOrDefault();

                    if (dataEdit.IsActive == "Y")
                    {
                        dataEdit.Rule_setter = AccountController.CurrentUserInfo.Name;
                        dataEdit.Rule_setting_Date = DateTime.Now.Date;
                        dataEdit.Auditor = "";
                        dataEdit.Audit_Date = null;
                        dataEdit.Status = "0";
                        db.SaveChanges();
                        result.RETURN_FLAG = true;
                        result.DESCRIPTION = Message_Type.delete_Success_Wait_Audit.GetDescription();
                    }
                    else
                    {
                        result.RETURN_FLAG = false;
                        result.DESCRIPTION = Message_Type.already_IsActiveN.GetDescription();
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

        #region sendD60ToAudit
        public MSGReturnModel sendD60ToAudit(string parmID, string auditor)
        {
            MSGReturnModel result = new MSGReturnModel();

            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                try
                {
                    string[] arrayParmID = parmID.Split(',');
                    string tempParmID = "";
                    for (int i = 0; i < arrayParmID.Length; i++)
                    {
                        tempParmID = arrayParmID[i];

                        Bond_Rating_Parm oldData = db.Bond_Rating_Parm
                                                     .Where(x => x.Parm_ID.ToString() == tempParmID)
                                                     .FirstOrDefault();

                        oldData.Auditor = auditor;
                        oldData.Status = "1";
                    }

                    db.SaveChanges();
                    result.RETURN_FLAG = true;
                    result.DESCRIPTION = Message_Type.send_To_Audit_Success.GetDescription();
                }
                catch (Exception ex)
                {
                    result.RETURN_FLAG = false;
                    result.DESCRIPTION = Message_Type.send_To_Audit_Fail.GetDescription(null, ex.Message);
                }
            }

            return result;
        }
        #endregion

        #region D60Audit
        public MSGReturnModel D60Audit(string parmID, string status)
        {
            MSGReturnModel result = new MSGReturnModel();

            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                try
                {
                    string[] arrayParmID = parmID.Split(',');
                    string tempParmID = arrayParmID[0];
                    Bond_Rating_Parm oldData = new Bond_Rating_Parm();
                    for (int i = 0; i < arrayParmID.Length; i++)
                    {
                        tempParmID = arrayParmID[i];

                        oldData = db.Bond_Rating_Parm
                                    .Where(x => x.Parm_ID.ToString() == tempParmID)
                                    .FirstOrDefault();

                        oldData.Audit_Date = DateTime.Now.Date;
                        bool changeFlag = false;
                        if (oldData.Status == "1") //只判斷資料狀態為 1:待複核
                        {
                            //status == "2" => 2:複核完成
                            if (status == "2" && oldData.IsActive == "N")//原為失效 => 新增案例
                            {
                                oldData.IsActive = "Y";
                                changeFlag = true;
                            }
                            else if (status == "2" && oldData.IsActive == "Y")//原為有效 => 刪除(失效)案例
                                oldData.IsActive = "N";
                        }
                        //status == "3" => 3:複核退回 資料保持不變
                        oldData.Status = status;
                        if (changeFlag)
                        {
                            var _change = db.Bond_Rating_Parm.Where(x =>
                             x.Parm_ID != oldData.Parm_ID &&
                             x.Rating_Object == oldData.Rating_Object &&
                             x.Status == "2" &&
                             x.IsActive == "Y")
                             .Where(x=>x.Rating_Org_Area == oldData.Rating_Org_Area,!oldData.Rating_Org_Area.IsNullOrWhiteSpace());
                            foreach (var change in _change)
                            {
                                change.IsActive = "N";
                            }
                        }
                    }

                    db.SaveChanges();

                    //if (oldData.Status == "2" && oldData.IsActive == "Y")
                    //{
                    //    string sql = $@"UPDATE Bond_Rating_Parm
                    //                    SET IsActive = 'N'
                    //                    WHERE Parm_ID <> {oldData.Parm_ID}
                    //                      AND Rating_Object = '{oldData.Rating_Object}'
                    //                      AND Status = '2'
                    //                      AND IsActive = 'Y';";

                    //    db.Database.ExecuteSqlCommand(sql);
                    //}

                    result.RETURN_FLAG = true;
                    if(oldData.Status == "2")
                        result.DESCRIPTION = Message_Type.Audit_Success.GetDescription();
                    else if (oldData.Status == "3")
                        result.DESCRIPTION = Message_Type.Reject_Success.GetDescription();
                }
                catch (Exception ex)
                {
                    result.RETURN_FLAG = false;
                    result.DESCRIPTION = Message_Type.Audit_Fail.GetDescription(null, ex.Message);
                }
            }

            return result;
        }
        #endregion

        #region getD61All
        public Tuple<bool, List<D61ViewModel>> getD61All()
        {
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                if (db.Bond_Assessment_Parm.Any())
                {
                    var users = db.IFRS9_User.AsNoTracking().ToList();
                    return new Tuple<bool, List<D61ViewModel>>
                                (
                                    true,
                                    (
                                        from q in db.Bond_Assessment_Parm.AsNoTracking()
                                                    .AsEnumerable()
                                        select DbToD61ViewModel(q, users)
                                    ).ToList()
                                );
                }
            }

            return new Tuple<bool, List<D61ViewModel>>(true, new List<D61ViewModel>());
        }
        #endregion

        #region Db 組成 DbToD61ViewModel
        private D61ViewModel DbToD61ViewModel(Bond_Assessment_Parm data, List<IFRS9_User> users)
        {
            string ruleSetterName = "";
            string auditorName = "";

            string statusName = "";

            var UserData = users.Where(x => x.User_Account == data.Rule_setter).FirstOrDefault();
            if (UserData != null)
            {
                ruleSetterName = UserData.User_Name;
            }
            
            UserData = users.Where(x => x.User_Account == data.Auditor).FirstOrDefault();
            if (UserData != null)
            {
                auditorName = UserData.User_Name;
            }            

            ProcessStatusList psl = new ProcessStatusList();
            List<SelectOption> selectOptionStatus = psl.statusOption;
            for (int i = 0; i < selectOptionStatus.Count; i++)
            {
                if (selectOptionStatus[i].Value == data.Status)
                {
                    statusName = selectOptionStatus[i].Text;
                    break;
                }
            }

            return new D61ViewModel()
            {
                Id = data.Id.ToString(),
                Assessment_Stage = data.Assessment_Stage,
                Assessment_Kind = data.Assessment_Kind,
                Assessment_Sub_Kind = data.Assessment_Sub_Kind,
                Check_Item_Code = data.Check_Item_Code,
                Check_Item = data.Check_Item,
                Check_Item_Memo = data.Check_Item_Memo,
                Check_Symbol = data.Check_Symbol,
                Threshold = TypeTransfer.doubleNToString(data.Threshold),
                Pass_Value = TypeTransfer.doubleNToString(data.Pass_Value),
                Rule_setter = data.Rule_setter,
                Rule_setter_Name = ruleSetterName,
                Rule_setting_Date = TypeTransfer.dateTimeNToString(data.Rule_setting_Date),
                Auditor = data.Auditor,
                Auditor_Name = auditorName,
                Audit_Date = TypeTransfer.dateTimeNToString(data.Audit_Date),
                Change_Status = data.Change_Status,
                Status = data.Status,
                Status_Name = statusName,
                IsActive = data.IsActive == "Y" ? "生效" : "失效"
            };
        }
        #endregion

        #region getD61
        public Tuple<bool, List<D61ViewModel>> getD61(D61ViewModel dataModel)
        {
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                var query = db.Bond_Assessment_Parm.AsNoTracking()
                              .Where(x => x.Assessment_Stage == dataModel.Assessment_Stage, dataModel.Assessment_Stage.IsNullOrWhiteSpace() == false)
                              .Where(x => x.Assessment_Kind == dataModel.Assessment_Kind, dataModel.Assessment_Kind.IsNullOrWhiteSpace() == false)
                              .Where(x => x.Assessment_Sub_Kind == dataModel.Assessment_Sub_Kind, dataModel.Assessment_Sub_Kind.IsNullOrWhiteSpace() == false)
                              .Where(x => x.Check_Item_Code == dataModel.Check_Item_Code, dataModel.Check_Item_Code.IsNullOrWhiteSpace() == false)
                              .Where(x => x.Check_Item == dataModel.Check_Item, dataModel.Check_Item.IsNullOrWhiteSpace() == false)
                              .Where(x => x.Status == dataModel.Status, dataModel.Status.IsNullOrWhiteSpace() == false)
                              .Where(x => x.IsActive == dataModel.IsActive, !dataModel.IsActive.IsNullOrWhiteSpace()).ToList();
                var users = db.IFRS9_User.AsNoTracking().ToList();
                return new Tuple<bool, List<D61ViewModel>>(query.Any(),
                    query.Select(x => { return DbToD61ViewModel(x, users); }).ToList());            
            }
        }
        #endregion

        #region saveD61
        public MSGReturnModel saveD61(string actionType, D61ViewModel dataModel)
        {
            MSGReturnModel result = new MSGReturnModel();

            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                try
                {
                    if (actionType == "Add")
                    {
                        var query = db.Bond_Assessment_Parm
                                      .Where(x => x.Check_Item_Code == dataModel.Check_Item_Code)
                                      .FirstOrDefault();
                        if (query != null)
                        {
                            if (query.IsActive == "Y")
                            {
                                result.RETURN_FLAG = false;
                                result.DESCRIPTION = "資料重複：您輸入的 評估項目編號 已存在";
                                return result;
                            }
                            else
                            {
                                query.Assessment_Stage = dataModel.Assessment_Stage;
                                query.Assessment_Kind = dataModel.Assessment_Kind;
                                query.Assessment_Sub_Kind = dataModel.Assessment_Sub_Kind;
                                query.Check_Item_Code = dataModel.Check_Item_Code;
                                query.Check_Item = dataModel.Check_Item;
                                query.Check_Item_Memo = dataModel.Check_Item_Memo;
                                query.Check_Symbol = dataModel.Check_Symbol;
                                if (dataModel.Threshold.IsNullOrWhiteSpace() == false)
                                {
                                    query.Threshold = double.Parse(dataModel.Threshold);
                                }
                                if (dataModel.Pass_Value.IsNullOrWhiteSpace() == false)
                                {
                                    query.Pass_Value = double.Parse(dataModel.Pass_Value);
                                }
                                query.Rule_setter = AccountController.CurrentUserInfo.Name;
                                query.Rule_setting_Date = DateTime.Now.Date;
                                query.Status = "0";
                                query.Change_Status = "I"; //新增
                                query.IsActive = "N"; //失效
                            }
                        }
                        else
                        {
                            Bond_Assessment_Parm dataAdd = new Bond_Assessment_Parm();

                            dataAdd.Assessment_Stage = dataModel.Assessment_Stage;
                            dataAdd.Assessment_Kind = dataModel.Assessment_Kind;
                            dataAdd.Assessment_Sub_Kind = dataModel.Assessment_Sub_Kind;
                            dataAdd.Check_Item_Code = dataModel.Check_Item_Code;
                            dataAdd.Check_Item = dataModel.Check_Item;
                            dataAdd.Check_Item_Memo = dataModel.Check_Item_Memo;
                            dataAdd.Check_Symbol = dataModel.Check_Symbol;
                            if (dataModel.Threshold.IsNullOrWhiteSpace() == false)
                            {
                                dataAdd.Threshold = double.Parse(dataModel.Threshold);
                            }
                            if (dataModel.Pass_Value.IsNullOrWhiteSpace() == false)
                            {
                                dataAdd.Pass_Value = double.Parse(dataModel.Pass_Value);
                            }
                            dataAdd.Rule_setter = AccountController.CurrentUserInfo.Name;
                            dataAdd.Rule_setting_Date = DateTime.Now.Date;
                            dataAdd.Status = "0";
                            dataAdd.IsActive = "N"; //失效
                            dataAdd.Change_Status = "I"; //新增
                            db.Bond_Assessment_Parm.Add(dataAdd);
                        }                
                    }
                    else if (actionType == "Modify")
                    {
                        var Id = Convert.ToInt32(dataModel.Id);
                        Bond_Assessment_Parm oldData = db.Bond_Assessment_Parm
                                                         .Where(x => x.Id == Id)
                                                         .FirstOrDefault();

                        Bond_Assessment_Parm AddData = new Bond_Assessment_Parm();
                        //會保留一筆生效,和一筆失效資料,但項目編號是相同
                        AddData.Assessment_Stage = dataModel.Assessment_Stage;
                        AddData.Assessment_Kind = dataModel.Assessment_Kind;
                        AddData.Assessment_Sub_Kind = dataModel.Assessment_Sub_Kind;
                        AddData.Check_Item_Code = dataModel.Check_Item_Code;
                        AddData.Check_Item = dataModel.Check_Item;
                        AddData.Check_Item_Memo = dataModel.Check_Item_Memo;
                        AddData.Check_Symbol = dataModel.Check_Symbol;
                        if (dataModel.Threshold.IsNullOrWhiteSpace() == false)
                        {
                            AddData.Threshold = double.Parse(dataModel.Threshold);
                        }
                        if (dataModel.Pass_Value.IsNullOrWhiteSpace() == false)
                        {
                            AddData.Pass_Value = double.Parse(dataModel.Pass_Value);
                        }
                        AddData.Rule_setter = AccountController.CurrentUserInfo.Name;
                        AddData.Rule_setting_Date = DateTime.Now.Date;
                        AddData.Auditor = "";
                        AddData.Audit_Date = null;
                        AddData.Status = "0";
                        AddData.Change_Status = "U"; //修改
                        AddData.IsActive = "N";
                        db.Bond_Assessment_Parm.Add(AddData);
                    }

                    db.SaveChanges();

                    result.RETURN_FLAG = true;
                }
                catch (DbUpdateException ex)
                {
                    result.RETURN_FLAG = false;
                    result.DESCRIPTION = ex.Message;
                }
            }

            return result;
        }
        #endregion

        #region deleteD61
        public MSGReturnModel deleteD61(string checkItemCode, string Id)
        {
            MSGReturnModel result = new MSGReturnModel();

            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                try
                {
                    var _Id = Convert.ToInt32(Id);
                    var query = db.Bond_Assessment_Parm
                                  .FirstOrDefault(x => x.Id == _Id);

                    if (query != null)
                    {
                        //if (query.IsActive == "N" && query.Change_Status != "U")
                        //{
                        //    result.DESCRIPTION = "目前為失效狀態,無須再刪除";
                        //}
                        //else
                        if (query.Change_Status == "D")
                        {
                            result.DESCRIPTION = "已經是刪除狀態等待複核";
                        }
                        else 
                        {
                            query.Rule_setter = AccountController.CurrentUserInfo.Name;
                            query.Rule_setting_Date = DateTime.Now.Date;
                            query.Change_Status = "D";
                            query.Status = "0";
                            query.Auditor = "";
                            query.Audit_Date = null;

                            db.SaveChanges(); //Save

                            result.RETURN_FLAG = true;
                            result.DESCRIPTION = Message_Type.delete_Success_Wait_Audit.GetDescription();
                        }
                    }
                    else
                    {
                        result.DESCRIPTION = Message_Type.not_Find_Any.GetDescription();
                    }
                }
                catch (DbUpdateException ex)
                {
                    result.RETURN_FLAG = false;
                    result.DESCRIPTION = ex.Message;
                }
            }

            return result;
        }
        #endregion

        #region sendD61ToAudit
        public MSGReturnModel sendD61ToAudit(List<D61ViewModel> dataModelList)
        {
            MSGReturnModel result = new MSGReturnModel();

            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                try
                {
                    int Id = 0;
                    for (int i = 0; i < dataModelList.Count; i++)
                    {
                        Id = Convert.ToInt32(dataModelList[i].Id);
                        Bond_Assessment_Parm oldData = db.Bond_Assessment_Parm
                                                         .Where(x => x.Id == Id)
                                                         .FirstOrDefault();
                        oldData.Auditor = dataModelList[i].Auditor;
                        oldData.Status = "1";
                    }

                    db.SaveChanges();
                    result.RETURN_FLAG = true;
                }
                catch (DbUpdateException ex)
                {
                    result.RETURN_FLAG = false;
                    result.DESCRIPTION = ex.Message;
                }
            }

            return result;
        }
        #endregion

        #region D61Audit
        public MSGReturnModel D61Audit(List<D61ViewModel> dataModelList)
        {
            MSGReturnModel result = new MSGReturnModel();

            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                try
                {
                    string checkItemCode = "";
                    int Id = 0;
                    for (int i = 0; i < dataModelList.Count; i++)
                    {
                        Id = Convert.ToInt32(dataModelList[i].Id);
                        Bond_Assessment_Parm oldData = db.Bond_Assessment_Parm.Where(x => x.Id == Id).FirstOrDefault();
                        checkItemCode = oldData.Check_Item_Code;
                        oldData.Auditor = AccountController.CurrentUserInfo.Name;
                        oldData.Audit_Date = DateTime.Now.Date;
                        oldData.Status = dataModelList[i].Status;
                        if (oldData.Status == "2") //狀態為覆核時才需要調整生效 by mark 2018/08/01
                        {
                            switch (oldData.Change_Status)
                            {
                                case "I":
                                    oldData.IsActive = "Y";
                                    break;
                                case "D":
                                    oldData.IsActive = "N";
                                    break;
                                case "U":
                                    oldData.IsActive = "Y";
                                        db.Bond_Assessment_Parm.
                                        Where(x => x.Check_Item_Code == checkItemCode &&
                                        x.Id != Id).ToList()
                                        .ForEach(x =>
                                        {
                                            x.IsActive = "N";
                                        });
                                    break;
                            }
                        }
                    }

                    db.SaveChanges();
                    result.RETURN_FLAG = true;
                }
                catch (DbUpdateException ex)
                {
                    result.RETURN_FLAG = false;
                    result.DESCRIPTION = ex.Message;
                }
            }

            return result;
        }
        #endregion

        #region getD68All
        public Tuple<bool, List<D68ViewModel>> getD68All()
        {
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                if (db.Risk_Parm.Any())
                {
                    var users = db.IFRS9_User.AsNoTracking().ToList();
                    return new Tuple<bool, List<D68ViewModel>>
                                (
                                    true,
                                    (
                                        from q in db.Risk_Parm.AsNoTracking()
                                                    .AsEnumerable()
                                        select DbToD68ViewModel(q, users)
                                    ).ToList()
                                );
                }
            }

            return new Tuple<bool, List<D68ViewModel>>(true, new List<D68ViewModel>());
        }
        #endregion

        #region Db 組成 DbToD68ViewModel
        private D68ViewModel DbToD68ViewModel(Risk_Parm data,List<IFRS9_User> users)
        {
            string includingIndName = "";
            string applyRangeName = "";
            string ruleSetterName = "";
            string auditorName = "";
            string statusName = "";
            string isActiveName = "";

            if (data.Including_Ind == "Y")
            {
                includingIndName = "是";
            }
            else if (data.Including_Ind == "N")
            {
                includingIndName = "否";
            }

            if (data.Apply_Range == "1")
            {
                applyRangeName = "1：以上";
            }
            else if (data.Apply_Range == "0")
            {
                applyRangeName = "0：以下";
            }

            ProcessStatusList psl = new ProcessStatusList();
            List<SelectOption> selectOptionStatus = psl.statusOption;
            for (int i = 0; i < selectOptionStatus.Count; i++)
            {
                if (selectOptionStatus[i].Value == data.Status)
                {
                    statusName = selectOptionStatus[i].Text;
                    break;
                }
            }

            IsActiveList ial = new IsActiveList();
            selectOptionStatus = ial.isActiveOption;
            for (int i = 0; i < selectOptionStatus.Count; i++)
            {
                if (selectOptionStatus[i].Value == data.IsActive)
                {
                    isActiveName = selectOptionStatus[i].Text;
                    break;
                }
            }

            var UserData = users.Where(x => x.User_Account == data.Rule_setter).FirstOrDefault();
            if (UserData != null)
            {
                ruleSetterName = UserData.User_Name;
            }

            UserData = users.Where(x => x.User_Account == data.Auditor).FirstOrDefault();
            if (UserData != null)
            {
                auditorName = UserData.User_Name;
            }
                
            return new D68ViewModel()
            {
                Rule_ID = data.Rule_ID.ToString(),
                Bond_Type = data.Bond_Type?.Trim(),
                Rating_Floor = data.Rating_Floor?.Trim(),
                Including_Ind = data.Including_Ind?.Trim(),
                Including_Ind_Name = includingIndName?.Trim(),
                Apply_Range = data.Apply_Range?.Trim(),
                Apply_Range_Name = applyRangeName?.Trim(),
                Data_Year = data.Data_Year?.Trim(),
                PD_Grade = data.PD_Grade.ToString(),
                Grade_Adjust = data.Grade_Adjust.ToString(),
                Rule_setter = data.Rule_setter?.Trim(),
                Rule_setter_Name = ruleSetterName?.Trim(),
                Rule_setting_Date = data.Rule_setting_Date.ToString("yyyy/MM/dd"),
                Auditor = data.Auditor?.Trim(),
                Auditor_Name = auditorName?.Trim(),
                Audit_Date = TypeTransfer.dateTimeNToString(data.Audit_Date),
                Status = data.Status?.Trim(),
                Status_Name = statusName?.Trim(),
                Rule_desc = data.Rule_desc?.Trim(),
                IsActive = data.IsActive?.Trim(),
                IsActive_Name = isActiveName?.Trim()
            };
        }
        #endregion

        #region getD68
        public Tuple<bool, List<D68ViewModel>> getD68(D68ViewModel dataModel)
        {
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                if (db.Risk_Parm.Any())
                {
                    var query = db.Risk_Parm.AsNoTracking()
                                  .Where(x => x.Rule_ID.ToString() == dataModel.Rule_ID, dataModel.Rule_ID.IsNullOrWhiteSpace() == false)
                                  .Where(x => x.Status == dataModel.Status, dataModel.Status.IsNullOrWhiteSpace() == false)
                                  .Where(x => x.IsActive == dataModel.IsActive, dataModel.IsActive.IsNullOrWhiteSpace() == false);
                    var users = db.IFRS9_User.AsNoTracking().ToList();
                    return new Tuple<bool, List<D68ViewModel>>(query.Any(),
                        query.AsEnumerable().Select(x => { return DbToD68ViewModel(x, users); }).ToList());
                }
            }

            return new Tuple<bool, List<D68ViewModel>>(false, new List<D68ViewModel>());
        }
        #endregion

        #region saveD68
        public MSGReturnModel saveD68(string actionType, D68ViewModel dataModel)
        {
            MSGReturnModel result = new MSGReturnModel();

            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                try
                {
                    Risk_Parm dataEdit = new Risk_Parm();

                    if (actionType == "Add")
                    {
                        dataEdit.IsActive = "N";
                    }
                    else if (actionType == "Modify")
                    {
                        dataEdit = db.Risk_Parm
                                     .Where(x => x.Rule_ID.ToString() == dataModel.Rule_ID)
                                     .FirstOrDefault();
                    }

                    dataEdit.Bond_Type = dataModel.Bond_Type;
                    dataEdit.Rating_Floor = dataModel.Rating_Floor;
                    dataEdit.Including_Ind = dataModel.Including_Ind;
                    dataEdit.Apply_Range = dataModel.Apply_Range;
                    dataEdit.Data_Year = dataModel.Data_Year;
                    dataEdit.PD_Grade = int.Parse(dataModel.PD_Grade);
                    dataEdit.Grade_Adjust = int.Parse(dataModel.Grade_Adjust);
                    dataEdit.Rule_setter = AccountController.CurrentUserInfo.Name;
                    dataEdit.Rule_setting_Date = DateTime.Now.Date;
                    dataEdit.Auditor = "";
                    dataEdit.Audit_Date = null;
                    dataEdit.Status = "0";
                    dataEdit.Rule_desc = dataModel.Rule_desc;

                    if (actionType == "Add")
                    {
                        db.Risk_Parm.Add(dataEdit);
                    }

                    db.SaveChanges();

                    result.RETURN_FLAG = true;
                    result.DESCRIPTION = Message_Type.insert_Success_Wait_Audit.GetDescription();
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

        #region Delete D68
        public MSGReturnModel deleteD68(string ruleID)
        {
            MSGReturnModel result = new MSGReturnModel();

            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                try
                {
                    Risk_Parm dataEdit = db.Risk_Parm
                                           .Where(x => x.Rule_ID.ToString() == ruleID)
                                           .FirstOrDefault();

                    if (dataEdit.IsActive == "Y")
                    {
                        dataEdit.Rule_setter = AccountController.CurrentUserInfo.Name;
                        dataEdit.Rule_setting_Date = DateTime.Now.Date;
                        dataEdit.Auditor = "";
                        dataEdit.Audit_Date = null;
                        dataEdit.Status = "0";

                        db.SaveChanges();
                        result.RETURN_FLAG = true;
                        result.DESCRIPTION = Message_Type.delete_Success_Wait_Audit.GetDescription();
                    }
                    else
                    {
                        result.RETURN_FLAG = false;
                        result.DESCRIPTION = Message_Type.already_IsActiveN.GetDescription();
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

        #region sendD68ToAudit
        public MSGReturnModel sendD68ToAudit(string ruleID, string auditor)
        {
            MSGReturnModel result = new MSGReturnModel();

            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                try
                {
                    string[] arrayRuleID = ruleID.Split(',');
                    string tempRuleID = "";
                    for (int i = 0; i < arrayRuleID.Length; i++)
                    {
                        tempRuleID = arrayRuleID[i];

                        Risk_Parm oldData = db.Risk_Parm
                                              .Where(x => x.Rule_ID.ToString() == tempRuleID)
                                              .FirstOrDefault();

                        oldData.Auditor = auditor;
                        oldData.Status = "1";
                    }

                    db.SaveChanges();
                    result.RETURN_FLAG = true;
                    result.DESCRIPTION = Message_Type.send_To_Audit_Success.GetDescription();
                }
                catch (Exception ex)
                {
                    result.RETURN_FLAG = false;
                    result.DESCRIPTION = Message_Type.send_To_Audit_Fail.GetDescription(null, ex.Message);
                }
            }

            return result;
        }
        #endregion

        #region D68Audit
        public MSGReturnModel D68Audit(string ruleID, string status)
        {
            MSGReturnModel result = new MSGReturnModel();

            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                try
                {
                    string[] arrayRuleID = ruleID.Split(',');
                    string tempRuleID = "";
                    for (int i = 0; i < arrayRuleID.Length; i++)
                    {
                        tempRuleID = arrayRuleID[i];

                        Risk_Parm oldData = db.Risk_Parm
                                              .Where(x => x.Rule_ID.ToString() == tempRuleID)
                                              .FirstOrDefault();

                        if (oldData.Status == "1")
                        {
                            //status == "2" => 2:複核完成
                            if (status == "2" && oldData.IsActive == "N")//原為失效 => 新增案例
                                oldData.IsActive = "Y";
                            else if (status == "2" && oldData.IsActive == "Y")//原為有效 => 刪除(失效)案例
                                oldData.IsActive = "N";
                            //status == "3" => 3:複核退回 資料保持不變
                        }

                        oldData.Audit_Date = DateTime.Now.Date;
                        oldData.Status = status;
                    }

                    db.SaveChanges();
                    result.RETURN_FLAG = true;
                    result.DESCRIPTION = Message_Type.Audit_Success.GetDescription();
                }
                catch (Exception ex)
                {
                    result.RETURN_FLAG = false;
                    result.DESCRIPTION = Message_Type.Audit_Fail.GetDescription(null, ex.Message);
                }
            }

            return result;
        }
        #endregion

        #region Get D69

        /// <summary>
        /// get D69 all data
        /// </summary>
        /// <returns></returns>
        public Tuple<bool, List<D69ViewModel>> getD69All()
        {
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                var users = db.IFRS9_User.AsNoTracking().ToList();
                if (db.Basic_Assessment_Parm.Any())
                {
                    return new Tuple<bool, List<D69ViewModel>>
                    (
                        true,
                        (
                            from q in db.Basic_Assessment_Parm.AsNoTracking()
                            .AsEnumerable().OrderBy(x => x.Rule_ID)
                            select DbToD69ViewModel(q, users)
                        ).ToList()
                    );
                }
            }

            return new Tuple<bool, List<D69ViewModel>>(true, new List<D69ViewModel>());
        }

        /// <summary>
        /// get D69 data
        /// </summary>
        /// <param name="dataModel"></param>
        /// <returns></returns>
        public Tuple<bool, List<D69ViewModel>> getD69(D69ViewModel dataModel)
        {
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                if (db.Basic_Assessment_Parm.Any())
                {
                    var users = db.IFRS9_User.AsNoTracking().ToList();
                    var query = from q in db.Basic_Assessment_Parm.AsNoTracking()
                                select q;

                    if (dataModel.Rule_ID.IsNullOrEmpty() == false)
                    {
                        query = query.Where(x => x.Rule_ID.ToString().Contains(dataModel.Rule_ID));
                    }

                    if (dataModel.Basic_Pass.IsNullOrEmpty() == false)
                    {
                        query = query.Where(x => x.Basic_Pass == dataModel.Basic_Pass);
                    }

                    if (dataModel.Rating_Ori_Good_Ind.IsNullOrEmpty() == false)
                    {
                        query = query.Where(x => x.Rating_Ori_Good_Ind == dataModel.Rating_Ori_Good_Ind);
                    }

                    if (dataModel.Status.IsNullOrEmpty() == false)
                    {
                        query = query.Where(x => x.Status == dataModel.Status);
                    }

                    if (dataModel.IsActive.IsNullOrEmpty() == false)
                    {
                        query = query.Where(x => x.IsActive == dataModel.IsActive);
                    }

                    return new Tuple<bool, List<D69ViewModel>>(query.Any(),
                        query.AsEnumerable().OrderBy(x => x.Rule_ID).Select(x => { return DbToD69ViewModel(x, users); }).ToList());
                }
            }

            return new Tuple<bool, List<D69ViewModel>>(false, new List<D69ViewModel>());
        }

        #endregion Get D69

        #region Db 組成 D69ViewModel

        /// <summary>
        /// Db 組成 D69ViewModel
        /// </summary>
        /// <param name="data"></param>
        /// <returns></returns>
        private D69ViewModel DbToD69ViewModel(Basic_Assessment_Parm data,List<IFRS9_User> users)
        {
            string basicPassName = "";
            string ratingOriGoodIndName = "";
            string includingIndName = "";
            string applyRangeName = "";
            string ratingCurrGoodIndName = "";
            string ruleSetterName = "";
            string auditorName = "";
            string statusName = "";
            string isActiveName = "";

            if (data.Basic_Pass == "Y")
            {
                basicPassName = "是";
            }
            else if (data.Basic_Pass == "N")
            {
                basicPassName = "否";
            }

            if (data.Rating_Ori_Good_Ind == "Y")
            {
                ratingOriGoodIndName = "是";
            }
            else if (data.Rating_Ori_Good_Ind == "N")
            {
                ratingOriGoodIndName = "否";
            }

            if (data.Including_Ind == "Y")
            {
                includingIndName = "是";
            }
            else if (data.Including_Ind == "N")
            {
                includingIndName = "否";
            }

            if (data.Apply_Range == "1")
            {
                applyRangeName = "1：以上";
            }
            else if (data.Apply_Range == "0")
            {
                applyRangeName = "0：以下";
            }

            if (data.Rating_Curr_Good_Ind == "Y")
            {
                ratingCurrGoodIndName = "是";
            }
            else if (data.Rating_Curr_Good_Ind == "N")
            {
                ratingCurrGoodIndName = "否";
            }

            ProcessStatusList psl = new ProcessStatusList();
            List<SelectOption> selectOptionStatus = psl.statusOption;
            for (int i = 0; i < selectOptionStatus.Count; i++)
            {
                if (selectOptionStatus[i].Value == data.Status)
                {
                    statusName = selectOptionStatus[i].Text;
                    break;
                }
            }

            IsActiveList ial = new IsActiveList();
            selectOptionStatus = ial.isActiveOption;
            for (int i = 0; i < selectOptionStatus.Count; i++)
            {
                if (selectOptionStatus[i].Value == data.IsActive)
                {
                    isActiveName = selectOptionStatus[i].Text;
                    break;
                }
            }

            var UserData = users.Where(x => x.User_Account == data.Rule_setter).FirstOrDefault();
            if (UserData != null)
            {
                ruleSetterName = UserData.User_Name;
            }

            UserData = users.Where(x => x.User_Account == data.Auditor).FirstOrDefault();
            if (UserData != null)
            {
                auditorName = UserData.User_Name;
            }

            return new D69ViewModel()
            {
                Rule_ID = data.Rule_ID.ToString(),
                Basic_Pass = data.Basic_Pass,
                Basic_Pass_Name = basicPassName,
                Rating_Ori_Good_Ind = data.Rating_Ori_Good_Ind,
                Rating_Ori_Good_Ind_Name = ratingOriGoodIndName,
                Rating_Notch = data.Rating_Notch.ToString(),
                Including_Ind = data.Including_Ind,
                Including_Ind_Name = includingIndName,
                Apply_Range = data.Apply_Range,
                Apply_Range_Name = applyRangeName,
                Rating_Curr_Good_Ind = data.Rating_Curr_Good_Ind,
                Rating_Curr_Good_Ind_Name = ratingCurrGoodIndName,
                Ori_Rating_Missing_Ind = data.Ori_Rating_Missing_Ind,
                Rule_setter = data.Rule_setter,
                Rule_setter_Name = ruleSetterName,
                Rule_setting_Date = TypeTransfer.dateTimeNToString(data.Rule_setting_Date),
                Auditor = data.Auditor,
                Auditor_Name = auditorName,
                Audit_Date = TypeTransfer.dateTimeNToString(data.Audit_Date),
                Status = data.Status,
                Status_Name = statusName,
                Rule_desc = data.Rule_desc,
                IsActive = data.IsActive,
                IsActive_Name = isActiveName
            };
        }

        #endregion Db 組成 D69ViewModel

        #region Save D69
        public MSGReturnModel saveD69(D69ViewModel dataModel)
        {
            MSGReturnModel result = new MSGReturnModel();

            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                try
                {
                    if (dataModel.ActionType == "Add")
                    {
                        Basic_Assessment_Parm data = new Basic_Assessment_Parm();

                        data.Basic_Pass = dataModel.Basic_Pass;
                        data.Rating_Ori_Good_Ind = dataModel.Rating_Ori_Good_Ind;
                        if (!dataModel.Rating_Notch.IsNullOrWhiteSpace())
                        {
                            data.Rating_Notch = int.Parse(dataModel.Rating_Notch);
                        }
                        data.Including_Ind = dataModel.Including_Ind;
                        data.Apply_Range = dataModel.Apply_Range;
                        data.Rating_Curr_Good_Ind = dataModel.Rating_Curr_Good_Ind;
                        data.Ori_Rating_Missing_Ind = dataModel.Ori_Rating_Missing_Ind;

                        data.Rule_setter = AccountController.CurrentUserInfo.Name;
                        data.Rule_setting_Date = DateTime.Now.Date;
                        data.Status = "0";
                        data.Rule_desc = dataModel.Rule_desc;
                        data.IsActive = "N";

                        db.Basic_Assessment_Parm.Add(data);
                    }
                    else if (dataModel.ActionType == "Modify")
                    {
                        Basic_Assessment_Parm oldData = db.Basic_Assessment_Parm.Where(x => x.Rule_ID.ToString() == dataModel.Rule_ID).FirstOrDefault();
                        oldData.Basic_Pass = dataModel.Basic_Pass;
                        oldData.Rating_Ori_Good_Ind = dataModel.Rating_Ori_Good_Ind;
                        oldData.Rating_Notch = null;
                        if (!dataModel.Rating_Notch.IsNullOrWhiteSpace())
                        {
                            oldData.Rating_Notch = int.Parse(dataModel.Rating_Notch);
                        }
                        oldData.Including_Ind = dataModel.Including_Ind;
                        oldData.Apply_Range = dataModel.Apply_Range;
                        oldData.Rating_Curr_Good_Ind = dataModel.Rating_Curr_Good_Ind;
                        oldData.Ori_Rating_Missing_Ind = dataModel.Ori_Rating_Missing_Ind;
                        oldData.Rule_setter = AccountController.CurrentUserInfo.Name;
                        oldData.Rule_setting_Date = DateTime.Now.Date;
                        oldData.Auditor = "";
                        oldData.Audit_Date = null;
                        oldData.Status = "0";
                        oldData.Rule_desc = dataModel.Rule_desc;
                    }

                    db.SaveChanges();

                    result.RETURN_FLAG = true;
                    result.DESCRIPTION = Message_Type.insert_Success_Wait_Audit.GetDescription();
                }
                catch (Exception ex)
                {
                    result.RETURN_FLAG = false;
                    result.DESCRIPTION = Message_Type.save_Fail.GetDescription(null, ex.Message);
                }
            }

            return result;
        }
        #endregion

        #region Delete D69
        public MSGReturnModel deleteD69(string ruleID)
        {
            MSGReturnModel result = new MSGReturnModel();

            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                try
                {
                    Basic_Assessment_Parm oldData = db.Basic_Assessment_Parm
                                                      .Where(x => x.Rule_ID.ToString() == ruleID)
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
                        result.RETURN_FLAG = false;
                        result.DESCRIPTION = Message_Type.already_IsActiveN.GetDescription();
                    }
                }
                catch (Exception ex)
                {
                    result.RETURN_FLAG = false;
                    result.DESCRIPTION = Message_Type.save_Fail.GetDescription(null, ex.Message);
                }
            }

            return result;
        }
        #endregion

        #region sendD69ToAudit

        /// <summary>
        /// send D69 to audit
        /// </summary>
        /// <param name="dataModelList"></param>
        /// <returns></returns>
        public MSGReturnModel sendD69ToAudit(List<D69ViewModel> dataModelList)
        {
            MSGReturnModel result = new MSGReturnModel();

            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                try
                {
                    int ruleID = 0;
                    for (int i = 0; i < dataModelList.Count; i++)
                    {
                        ruleID = int.Parse(dataModelList[i].Rule_ID);
                        Basic_Assessment_Parm oldData = db.Basic_Assessment_Parm.Where(x => x.Rule_ID == ruleID).FirstOrDefault();
                        oldData.Auditor = dataModelList[i].Auditor;
                        oldData.Status = "1";
                    }

                    db.SaveChanges();
                    result.RETURN_FLAG = true;
                    result.DESCRIPTION = Message_Type.send_To_Audit_Success.GetDescription();
                }
                catch (DbUpdateException ex)
                {
                    result.RETURN_FLAG = false;
                    result.DESCRIPTION = Message_Type.send_To_Audit_Fail.GetDescription(null, ex.Message);
                }
            }

            return result;
        }
        #endregion

        #region D69Audit
        public MSGReturnModel D69Audit(List<D69ViewModel> dataModelList)
        {
            MSGReturnModel result = new MSGReturnModel();

            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                try
                {
                    int ruleID = 0;

                    for (int i = 0; i < dataModelList.Count; i++)
                    {
                        ruleID = int.Parse(dataModelList[i].Rule_ID);

                        Basic_Assessment_Parm oldData = db.Basic_Assessment_Parm.Where(x => x.Rule_ID == ruleID).FirstOrDefault();
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

                    db.SaveChanges();
                    result.RETURN_FLAG = true;
                    result.DESCRIPTION = Message_Type.Audit_Success.GetDescription("D69");
                }
                catch (DbUpdateException ex)
                {
                    result.RETURN_FLAG = false;
                    result.DESCRIPTION = Message_Type.Audit_Fail.GetDescription("D69", ex.Message); 
                }
            }

            return result;
        }

        #endregion

        #region getD62
        public Tuple<string, List<D62ViewModel>> getD62(string reportDateStart, string reportDateEnd,
                                                        string referenceNbr, string bondNumber,
                                                        string basicPass,
                                                        string watchIND, string warningIND,
                                                        string chgInSpreadIND, string beforeHasChgInSpread)
        {
            List<D62ViewModel> D62 = new List<D62ViewModel>();
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                if (db.Bond_Basic_Assessment.Any())
                {
                    var data = from q in db.Bond_Basic_Assessment.AsNoTracking()
                               select q;

                    DateTime dateReportDateStart = DateTime.Now.Date;
                    DateTime dateReportDateEnd = DateTime.Now.Date;

                    if (reportDateStart.IsNullOrWhiteSpace() == false)
                    {
                        dateReportDateStart = DateTime.Parse(reportDateStart);
                        data = data.Where(x => x.Report_Date >= dateReportDateStart);
                    }

                    if (reportDateEnd.IsNullOrWhiteSpace() == false)
                    {
                        dateReportDateEnd = DateTime.Parse(reportDateEnd);
                        data = data.Where(x => x.Report_Date <= dateReportDateEnd);
                    }

                    data = data.Where(x => x.Reference_Nbr == referenceNbr, referenceNbr.IsNullOrWhiteSpace() == false)
                               .Where(x => x.Bond_Number == bondNumber, bondNumber.IsNullOrWhiteSpace() == false)
                               .Where(x => x.Basic_Pass == basicPass, basicPass.IsNullOrWhiteSpace() == false)
                               .Where(x => x.Watch_IND == watchIND, watchIND.IsNullOrWhiteSpace() == false)
                               .Where(x => x.Warning_IND == warningIND, warningIND.IsNullOrWhiteSpace() == false)
                               .Where(x => x.Chg_In_Spread.ToString() != "", chgInSpreadIND == "Y")
                               .Where(x => x.Chg_In_Spread.ToString() == "", chgInSpreadIND == "N");

                    List<Bond_Basic_Assessment> dataD62 = data.ToList();

                    List<Bond_Basic_Assessment> D62Before = db.Bond_Basic_Assessment.AsNoTracking()
                                                              .Where(x => x.Report_Date >= dateReportDateStart
                                                                       && x.Report_Date <= dateReportDateEnd).ToList();
                    if (beforeHasChgInSpread == "Y")
                    {
                        D62Before = D62Before.Where(x => x.Chg_In_Spread.ToString() != "").ToList();

                        for (int i = 0; i < dataD62.Count; i++)
                        {
                            var oneD62Before =  D62Before.Where(x => x.Report_Date < dataD62[i].Report_Date
                                                                  && x.Bond_Number == dataD62[i].Bond_Number
                                                                  && x.Lots == dataD62[i].Lots
                                                                  && x.Portfolio_Name == dataD62[i].Portfolio_Name)
                                                        .OrderByDescending(x => x.Report_Date)
                                                        .FirstOrDefault();
                            if (oneD62Before != null)
                            {
                                dataD62.Add(oneD62Before);
                            }
                        }
                    }

                    D62 = dataD62.Distinct()
                          .OrderBy(x=> x.Bond_Number).OrderBy(x => x.Lots).OrderBy(x => x.Portfolio_Name).OrderByDescending(x => x.Report_Date)
                          .AsEnumerable()
                          .Select(x => {return DbToD62Model(x);}).ToList();

                    List <IFRS9_User> ifrs9UserList = db.IFRS9_User.AsNoTracking().ToList();

                    for (int i=0; i < D62.Count;i++)
                    {
                        string Watch_IND_Override_User = D62[i].Watch_IND_Override_User;
                        string Warning_IND_Override_User = D62[i].Warning_IND_Override_User;

                        var ifrs9User = ifrs9UserList.Where(x => x.User_Account == Watch_IND_Override_User)
                                                     .FirstOrDefault();
                        if (ifrs9User != null)
                        {
                            D62[i].Watch_IND_Override_User_Name = ifrs9User.User_Name;
                        }

                        ifrs9User = ifrs9UserList.Where(x => x.User_Account == Warning_IND_Override_User)
                                                 .FirstOrDefault();
                        if (ifrs9User != null)
                        {
                            D62[i].Warning_IND_Override_User_Name = ifrs9User.User_Name;
                        }
                    }
                }
            }

            string message = "查無資料";

            if (D62.Any())
            {
                message = getD62Message(reportDateStart,reportDateEnd);
            }

            return new Tuple<string, List<D62ViewModel>>(message, D62);
        }

        #endregion

        #region Db 組成 D62ViewModel
        private D62ViewModel DbToD62Model(Bond_Basic_Assessment data)
        {
            string Watch_IND_Override_User_Name = "";
            string Warning_IND_Override_User_Name = "";

            return new D62ViewModel()
            {
                Reference_Nbr = data.Reference_Nbr,
                Version = data.Version.ToString(),
                Report_Date = data.Report_Date.ToString("yyyy/MM/dd"),
                Processing_Date = data.Processing_Date.ToString("yyyy/MM/dd"),
                Bond_Type = data.Bond_Type,
                Assessment_Sub_Kind = data.Assessment_Sub_Kind,
                Lien_position = data.Lien_position,
                Bond_Number = data.Bond_Number,
                Lots = data.Lots,
                Portfolio = data.Portfolio,
                Portfolio_Name = data.Portfolio_Name,
                Origination_Date = TypeTransfer.dateTimeNToString(data.Origination_Date),
                Original_External_Rating = data.Original_External_Rating.ToString(),
                Current_External_Rating = data.Current_External_Rating.ToString(),
                Rating_Ori_Good_IND = data.Rating_Ori_Good_IND,
                Rating_Curr_Good_Ind = data.Rating_Curr_Good_Ind,
                Curr_Ori_Rating_Diff = data.Curr_Ori_Rating_Diff.ToString(),
                Basic_Pass = data.Basic_Pass,
                Map_Rule_Id_D69 = data.Map_Rule_Id_D69.ToString(),
                Cost_Value = data.Cost_Value.ToString(),
                Market_Value_Ori = data.Market_Value_Ori.ToString(),
                Value_Change_Ratio = data.Value_Change_Ratio.ToString(),
                Value_Change_Ratio_Pass = data.Value_Change_Ratio_Pass,
                Accumulation_Loss_last_Month = data.Accumulation_Loss_last_Month.ToString(),
                Accumulation_Loss_This_Month = data.Accumulation_Loss_This_Month.ToString(),
                Watch_IND = data.Watch_IND,
                Map_Rule_Id_D70 = data.Map_Rule_Id_D70,
                Warning_IND = data.Warning_IND,
                Map_Rule_Id_D71 = data.Map_Rule_Id_D71,
                Chg_In_Spread = data.Chg_In_Spread.ToString(),
                Spread_Change_Over = data.Spread_Change_Over,
                Chg_In_Spread_last_Month = data.Chg_In_Spread_last_Month.ToString(),
                Chg_In_Spread_This_Month = data.Chg_In_Spread_This_Month.ToString(),
                Watch_IND_Override = data.Watch_IND_Override,
                Watch_IND_Override_Desc = data.Watch_IND_Override_Desc,
                Watch_IND_Override_Date = TypeTransfer.dateTimeNToString(data.Watch_IND_Override_Date),
                Watch_IND_Override_User = data.Watch_IND_Override_User,
                Watch_IND_Override_User_Name = Watch_IND_Override_User_Name,
                Warning_IND_Override = data.Warning_IND_Override,
                Warning_IND_Override_DESC = data.Warning_IND_Override_DESC,
                Warning_IND_Override_Date = TypeTransfer.dateTimeNToString(data.Warning_IND_Override_Date),
                Warning_IND_Override_User = data.Warning_IND_Override_User,
                Warning_IND_Override_User_Name = Warning_IND_Override_User_Name
            };
        }

        #endregion

        #region getD62Message
        public string getD62Message(string reportDateStart, string reportDateEnd)
        {
            string message = "";

            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                DateTime dateReportDateStart = DateTime.Parse(reportDateStart);
                DateTime dateReportDateEnd = DateTime.Parse(reportDateEnd);

                int basicPassYCount = 0;
                int basicPassNCount = 0;
                int basicPassNULLCount = 0;
                int watchINDYCount = 0;
                int watchINDNCount = 0;
                int watchINDNULLCount = 0;

                List<Bond_Basic_Assessment> listBond_Basic_Assessment = db.Bond_Basic_Assessment.AsNoTracking()
                                                                          .Where(x => x.Report_Date >= dateReportDateStart
                                                                                   && x.Report_Date <= dateReportDateEnd)
                                                                          .ToList();

                basicPassYCount = listBond_Basic_Assessment.Where(x => x.Basic_Pass == "Y").Count();
                basicPassNCount = listBond_Basic_Assessment.Where(x => x.Basic_Pass == "N").Count();
                basicPassNULLCount = listBond_Basic_Assessment.Where(x => x.Basic_Pass.ToString() == "").Count();

                watchINDYCount = listBond_Basic_Assessment.Where(x => x.Watch_IND == "Y").Count();
                watchINDNCount = listBond_Basic_Assessment.Where(x => x.Watch_IND == "N").Count();
                watchINDNULLCount = listBond_Basic_Assessment.Where(x => x.Watch_IND == "" || x.Watch_IND == null).Count();

                message = $"基準日：{reportDateStart} ~ {reportDateEnd}，<br/>";
                message += $@"基本要件通過=Y：{basicPassYCount.ToString()} 筆，<br/>基本要件通過=N：{basicPassNCount.ToString()} 筆，<br/>基本要件通過=空：{basicPassNULLCount.ToString()} 筆。<br/>";
                message += $@"觀察名單=Y：{watchINDYCount.ToString()} 筆，<br/>觀察名單=N：{watchINDNCount.ToString()} 筆，<br/>觀察名單=空：{watchINDNULLCount.ToString()} 筆。";
            }

            return message;
        }
        #endregion

        #region DownLoadD62Excel
        public MSGReturnModel DownLoadD62Excel(string path, List<D62ViewModel> dbDatas)
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            result.DESCRIPTION = Message_Type.download_Fail
                                 .GetDescription("D62", Message_Type.not_Find_Any.GetDescription());

            if (dbDatas.Any())
            {
                DataTable datas = getD62ModelFromDb(dbDatas).Item1;

                result.DESCRIPTION = FileRelated.DataTableToExcel(datas, path, Excel_DownloadName.D62);
                result.RETURN_FLAG = string.IsNullOrWhiteSpace(result.DESCRIPTION);

                if (result.RETURN_FLAG)
                {
                    result.DESCRIPTION = Message_Type.download_Success.GetDescription();
                }
            }

            return result;
        }

        #endregion

        #region getD62ModelFromDb
        private Tuple<DataTable> getD62ModelFromDb(List<D62ViewModel> dbDatas)
        {
            DataTable dt = new DataTable();

            try
            {
                dt.Columns.Add("帳戶編號/群組編號", typeof(object));
                dt.Columns.Add("資料版本", typeof(object));
                dt.Columns.Add("評估基準日/報導日", typeof(object));
                dt.Columns.Add("資料處理日期", typeof(object));
                dt.Columns.Add("債券種類", typeof(object));
                dt.Columns.Add("Stage評估次分類", typeof(object));
                dt.Columns.Add("債券擔保順位", typeof(object));
                dt.Columns.Add("債券編號", typeof(object));
                dt.Columns.Add("Lots", typeof(object));
                dt.Columns.Add("Portfolio", typeof(object));
                dt.Columns.Add("Portfolio英文", typeof(object));
                dt.Columns.Add("債券購入(認列)日期", typeof(object));
                dt.Columns.Add("起始(原始購買)外部評等", typeof(object));
                dt.Columns.Add("最近(報導日)外部信評評等", typeof(object));
                dt.Columns.Add("原始購買評等是否符合信用風險低條件", typeof(object));
                dt.Columns.Add("報導日是否符合信用風險低條件", typeof(object));
                dt.Columns.Add("評等下降數", typeof(object));
                dt.Columns.Add("基本要件通過與否", typeof(object));
                dt.Columns.Add("對應D69-基本要件參數檔規則編號", typeof(object));
                dt.Columns.Add("金融資產餘額-攤銷後之成本數(原幣)", typeof(object));
                dt.Columns.Add("市價(原幣)", typeof(object));
                dt.Columns.Add("未實現累計損失率", typeof(object));
                dt.Columns.Add("未實現累計損失率是否低於30%", typeof(object));
                dt.Columns.Add("未實現累計損失月數_上個月狀況", typeof(object));
                dt.Columns.Add("未實現累計損失月數_本月狀況", typeof(object));
                dt.Columns.Add("信用利差", typeof(object));
                dt.Columns.Add("信用利差超過拓寬300基點(含)以上", typeof(object));
                dt.Columns.Add("信用利差_上個月狀況", typeof(object));
                dt.Columns.Add("信用利差_本月狀況(信用利差目前累計逾越月數)", typeof(object));
                dt.Columns.Add("是否為觀察名單", typeof(object));
                dt.Columns.Add("對應D70-觀察名單參數檔規則編號", typeof(object));
                dt.Columns.Add("是否為預警名單", typeof(object));
                dt.Columns.Add("對應D71-預警名單參數檔規則編號", typeof(object));
                dt.Columns.Add("觀察名單手動調整狀態", typeof(object));
                dt.Columns.Add("觀察名單手動調整原因", typeof(object));
                dt.Columns.Add("觀察名單手動調整日期", typeof(object));
                dt.Columns.Add("觀察名單手動調整者帳號", typeof(object));
                dt.Columns.Add("觀察名單手動調整者名稱", typeof(object));
                dt.Columns.Add("預警名單手動調整狀態", typeof(object));
                dt.Columns.Add("預警名單手動調整原因", typeof(object));
                dt.Columns.Add("預警名單手動調整日期", typeof(object));
                dt.Columns.Add("預警名單手動調整者帳號", typeof(object));
                dt.Columns.Add("預警名單手動調整者名稱", typeof(object));

                foreach (D62ViewModel item in dbDatas)
                {
                    var nrow = dt.NewRow();

                    nrow["帳戶編號/群組編號"] = item.Reference_Nbr;
                    nrow["資料版本"] = item.Version;
                    nrow["評估基準日/報導日"] = item.Report_Date;
                    nrow["資料處理日期"] = item.Processing_Date;
                    nrow["債券種類"] = item.Bond_Type;
                    nrow["Stage評估次分類"] = item.Assessment_Sub_Kind;
                    nrow["債券擔保順位"] = item.Lien_position;
                    nrow["債券編號"] = item.Bond_Number;
                    nrow["Lots"] = item.Lots;
                    nrow["Portfolio"] = item.Portfolio;
                    nrow["Portfolio英文"] = item.Portfolio_Name;
                    nrow["債券購入(認列)日期"] = item.Origination_Date;
                    nrow["起始(原始購買)外部評等"] = item.Original_External_Rating;
                    nrow["最近(報導日)外部信評評等"] = item.Current_External_Rating;
                    nrow["原始購買評等是否符合信用風險低條件"] = item.Rating_Ori_Good_IND;
                    nrow["報導日是否符合信用風險低條件"] = item.Rating_Curr_Good_Ind;
                    nrow["評等下降數"] = item.Curr_Ori_Rating_Diff;
                    nrow["基本要件通過與否"] = item.Basic_Pass;
                    nrow["對應D69-基本要件參數檔規則編號"] = item.Map_Rule_Id_D69;
                    nrow["金融資產餘額-攤銷後之成本數(原幣)"] = item.Cost_Value;
                    nrow["市價(原幣)"] = item.Market_Value_Ori;
                    nrow["未實現累計損失率"] = item.Value_Change_Ratio;
                    nrow["未實現累計損失率是否低於30%"] = item.Value_Change_Ratio_Pass;
                    nrow["未實現累計損失月數_上個月狀況"] = item.Accumulation_Loss_last_Month;
                    nrow["未實現累計損失月數_本月狀況"] = item.Accumulation_Loss_This_Month;
                    nrow["信用利差"] = item.Chg_In_Spread;
                    nrow["信用利差超過拓寬300基點(含)以上"] = item.Spread_Change_Over;
                    nrow["信用利差_上個月狀況"] = item.Chg_In_Spread_last_Month;
                    nrow["信用利差_本月狀況(信用利差目前累計逾越月數)"] = item.Chg_In_Spread_This_Month;
                    nrow["是否為觀察名單"] = item.Watch_IND;
                    nrow["對應D70-觀察名單參數檔規則編號"] = item.Map_Rule_Id_D70;
                    nrow["是否為預警名單"] = item.Warning_IND;
                    nrow["對應D71-預警名單參數檔規則編號"] = item.Map_Rule_Id_D71;
                    nrow["觀察名單手動調整狀態"] = item.Watch_IND_Override;
                    nrow["觀察名單手動調整原因"] = item.Watch_IND_Override_Desc;
                    nrow["觀察名單手動調整日期"] = item.Watch_IND_Override_Date;
                    nrow["觀察名單手動調整者帳號"] = item.Watch_IND_Override_User;
                    nrow["觀察名單手動調整者名稱"] = item.Watch_IND_Override_User_Name;
                    nrow["預警名單手動調整狀態"] = item.Warning_IND_Override;
                    nrow["預警名單手動調整原因"] = item.Warning_IND_Override_DESC;
                    nrow["預警名單手動調整日期"] = item.Warning_IND_Override_Date;
                    nrow["預警名單手動調整者帳號"] = item.Warning_IND_Override_User;
                    nrow["預警名單手動調整者名稱"] = item.Warning_IND_Override_User_Name;

                    dt.Rows.Add(nrow);
                }
            }
            catch
            {
            }

            return new Tuple<DataTable>(dt);
        }

        #endregion

        #region Save D62
        public MSGReturnModel saveD62(string reportDate, string version)
        {
            MSGReturnModel result = new MSGReturnModel();

            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                string tempReference_Nbr = "";

                try
                {
                    DateTime dateReportDate = DateTime.Parse(reportDate);

                    string reportDate2 = DateTime.Parse(reportDate).ToString("yyyy-MM-dd");
                    if (db.IFRS9_EL.AsNoTracking().Where(x => x.Report_Date == reportDate
                                                           || x.Report_Date == reportDate2).Any() == true)
                    {
                        result.RETURN_FLAG = false;
                        result.DESCRIPTION = $"{reportDate} 的資料已執行減損階段最終確認，無法再執行。";
                        return result;
                    }

                    var query = db.Bond_Account_Info.AsNoTracking()
                                  .Where(x => x.Report_Date == dateReportDate)
                                  .OrderByDescending(x => x.Version)
                                  .FirstOrDefault();

                    if (query == null)
                    {
                        result.RETURN_FLAG = false;
                        result.DESCRIPTION = $"A41：Bond_Account_Info 查無 {reportDate} 的資料";
                        return result;
                    }

                    int intVersion = TypeTransfer.intNToInt(query.Version);

                    var addData = db.Bond_Account_Info.AsNoTracking().Where(x => x.Report_Date == dateReportDate
                                                                              && x.Version == intVersion
                                                                              && x.IAS39_CATEGORY.StartsWith("FVPL") == false
                                                                              && x.Assessment_Check!="N") //190510 PG&E需求，排除A41註記為N的部位
                                                                     .ToList();
                    if (addData.Any() == false)
                    {
                        result.RETURN_FLAG = false;
                        result.DESCRIPTION = $"{reportDate}、版本 {intVersion.ToString()} 的資料，都是公報分類(IAS39_CATEGORY)中，開頭帶有「FVPL」字串的券";
                        return result;
                    }

                    if (addData.Where(x=>x.Principal.HasValue == false || x.Market_Value_Ori.HasValue == false).Any())
                    {
                        result.RETURN_FLAG = false;
                        result.DESCRIPTION = "A41：Bond_Account_Info 有成本價或是市價有空值的狀況";
                        return result;
                    }

                    var D68Data = db.Risk_Parm.AsNoTracking().Where(x => x.IsActive == "Y").ToList(); ;
                    if (!D68Data.Any())
                    {
                        result.RETURN_FLAG = false;
                        result.DESCRIPTION = "D68：信用風險低參數檔 沒有生效的規則";
                        return result;
                    }
                    var D69Data = db.Basic_Assessment_Parm.AsNoTracking().Where(x => x.IsActive == "Y").ToList();
                    if (!D69Data.Any())
                    {
                        result.RETURN_FLAG = false;
                        result.DESCRIPTION = "D69：基本要件參數檔 沒有生效的規則";
                        return result;
                    }

                    result = checkD70D71();
                    if (result.RETURN_FLAG == false)
                    {
                        return result;
                    }

                    result = new MSGReturnModel();

                    var B01Data = db.IFRS9_Main.AsNoTracking().Where(x=>x.Report_Date == dateReportDate).ToList();
                    if (!B01Data.Any())
                    {
                        result.RETURN_FLAG = false;
                        result.DESCRIPTION = "B01：IFRS9_Main 無資料，無法抓取相關值";
                        return result;
                    }

                    DateTime lastReportDate = dateReportDate.AddDays(-dateReportDate.Day);
                    DateTime lastTwoReportDate = dateReportDate.AddMonths(-1).AddDays(-dateReportDate.AddMonths(-1).Day);
                    List<DateTime> dts = new List<DateTime>() { lastReportDate, lastTwoReportDate };

                    #region 加入刪除 D63 & D65
                    var _cr = $"{ dateReportDate.ToString("yyyyMMdd")}_V{intVersion}";

                    var D64Files = db.Bond_Quantitative_Result_File
                        .Where(x => x.Check_Reference.StartsWith(_cr)).ToList();

                    var D66Files = db.Bond_Qualitative_Assessment_Result_File
                        .Where(x => x.Check_Reference.StartsWith(_cr)).ToList();

                    D64Files.ForEach(x =>
                    {
                        try
                        {
                            File.Delete(x.File_path);
                        }
                        catch
                        {

                        }
                    });

                    D66Files.ForEach(x =>
                    {
                        try
                        {
                            File.Delete(x.File_path);
                        }
                        catch
                        {

                        }
                    });

                    db.Bond_Quantitative_Result_File.RemoveRange(D64Files);
                    db.Bond_Qualitative_Assessment_Result_File.RemoveRange(D66Files);

                    string delSql = string.Empty;

                    delSql = $@"
DELETE Bond_Quantitative_Resource
WHERE Report_Date = @Report_Date
AND Version = @Version ;

DELETE Bond_Quantitative_Result
WHERE Report_Date = @Report_Date
AND Version = @Version ;

DELETE Bond_Qualitative_Assessment
WHERE Report_Date = @Report_Date
AND Version = @Version ;

DELETE Bond_Qualitative_Assessment_Result
WHERE Report_Date = @Report_Date
AND Version = @Version ; ";

                    try
                    {
                        int i = db.Database.ExecuteSqlCommand(delSql,
                            new List<SqlParameter>()
                            {
                                  new SqlParameter("Report_Date", dateReportDate.ToString("yyyy-MM-dd")),
                                  new SqlParameter("Version", intVersion.ToString())
                            }.ToArray());

                        db.SaveChanges();
                    }
                    catch
                    {

                    }

                    #endregion


                    db.Database.ExecuteSqlCommand(string.Format(@"DELETE FROM Bond_Basic_Assessment 
                                                                  WHERE Report_Date = '{0}'", reportDate));

                    StringBuilder sb = new StringBuilder();
                    List<IFRS9_Main> ifrs9MainList = B01Data;
                    List<Risk_Parm> riskParmList = D68Data;
                    List<Basic_Assessment_Parm> bapList = D69Data;
                    List<Bond_Basic_Assessment> bbaList = db.Bond_Basic_Assessment.AsNoTracking().Where(x=> dts.Contains(x.Report_Date)).ToList();
                    List<Watching_List_Parm> ListWatch = db.Watching_List_Parm.AsNoTracking().Where(x => x.IsActive == "Y").ToList();
                    List<Warning_List_Parm> ListWarning = db.Warning_List_Parm.AsNoTracking().Where(x => x.IsActive == "Y").ToList();
                    List<Bond_Spread_Info> ListA96 = db.Bond_Spread_Info.AsNoTracking().Where(x => x.Report_Date == dateReportDate).ToList();
                    for (int i = 0; i < addData.Count(); i++)
                    {
                        tempReference_Nbr = addData[i].Reference_Nbr;
                        string tempIAS39_CATEGORY = addData[i].IAS39_CATEGORY;
                        string tempVersion = addData[i].Version.ToString();
                        string tempReport_Date = TypeTransfer.dateTimeNToString(addData[i].Report_Date);
                        string tempProcessing_Date = DateTime.Now.ToString("yyyy/MM/dd");
                        string tempBond_Type = (addData[i].Bond_Type == null ? "" : addData[i].Bond_Type);
                        string tempAssessment_Sub_Kind = (addData[i].Assessment_Sub_Kind == null ? "" : addData[i].Assessment_Sub_Kind);
                        string tempLien_position = (addData[i].Lien_position == null ? "" : addData[i].Lien_position);
                        string tempBond_Number = addData[i].Bond_Number;
                        string tempLots = addData[i].Lots;
                        string tempPortfolio = addData[i].Portfolio;
                        string tempPortfolio_Name = addData[i].Portfolio_Name;
                        string tempOrigination_Date = TypeTransfer.dateTimeNToString(addData[i].Origination_Date);
                        string tempOriginal_External_Rating = "";
                        string tempLastTwoCurrentExternalRating = "";
                        string tempLastCurrentExternalRating = "";
                        string tempCurrent_External_Rating = "";
                        string tempRating_Ori_Good_IND = "";
                        string tempRating_Curr_Good_Ind = "";
                        string tempCurr_Ori_Rating_Diff = "";
                        string tempLastTwoBasicPass = "";
                        string tempLastBasicPass = "";
                        string tempBasic_Pass = "";
                        string tempMap_Rule_Id_D69 = "";
                        string tempCost_Value = addData[i].Principal.ToString();
                        string tempMarket_Value_Ori = addData[i].Market_Value_Ori.ToString();
                        string tempValue_Change_Ratio = "";
                        string tempValue_Change_Ratio_Pass = "";
                        string tempAccumulation_Loss_last_Month = "0";
                        string tempAccumulation_Loss_This_Month = "";
                        string tempWatch_IND = "";
                        string tempMap_Rule_Id_D70 = "";
                        string tempWarning_IND = "";
                        string tempMap_Rule_Id_D71 = "";
                        string tempChg_In_Spread = "";
                        string tempSpread_Change_Over = "";
                        string tempChg_In_Spread_last_Month = "";
                        string tempChg_In_Spread_This_Month = "";

                        IFRS9_Main ifrs9Main = ifrs9MainList.Where(x => x.Reference_Nbr == tempReference_Nbr)
                                                            .FirstOrDefault();
                        if (ifrs9Main != null)
                        {
                            tempOriginal_External_Rating = ifrs9Main.Original_External_Rating;
                            tempCurrent_External_Rating = ifrs9Main.Current_External_Rating;
                        }

                        tempOriginal_External_Rating = tempOriginal_External_Rating ?? "";
                        tempCurrent_External_Rating = tempCurrent_External_Rating ?? "";

                        Risk_Parm riskParm = riskParmList.Where(x => x.Bond_Type == tempBond_Type)
                                                         .FirstOrDefault();
                        if (riskParm != null)
                        {
                            if ((tempOriginal_External_Rating ?? "") != "")
                            {
                                tempRating_Ori_Good_IND = CompareRange(riskParm.Including_Ind, riskParm.Apply_Range, riskParm.Grade_Adjust, int.Parse(tempOriginal_External_Rating));
                            }

                            if ((tempCurrent_External_Rating ?? "") != "")
                            {
                                tempRating_Curr_Good_Ind = CompareRange(riskParm.Including_Ind, riskParm.Apply_Range, riskParm.Grade_Adjust, int.Parse(tempCurrent_External_Rating));
                            }
                        }

                        tempRating_Ori_Good_IND = tempRating_Ori_Good_IND ?? "";
                        tempRating_Curr_Good_Ind = tempRating_Curr_Good_Ind ?? "";

                        if ((tempOriginal_External_Rating ?? "") != "" && (tempCurrent_External_Rating ?? "") != "")
                        {
                            tempCurr_Ori_Rating_Diff = (int.Parse(tempCurrent_External_Rating) - int.Parse(tempOriginal_External_Rating)).ToString();
                        }

                        tempBasic_Pass = "Y";
                        foreach (Basic_Assessment_Parm itemBAP in bapList)
                        {
                            string Match_Rating_Ori_Good_Ind = "";
                            string Match_Rating_Notch = "";
                            string Match_Rating_Curr_Good_Ind = "";
                            string Match_Ori_Rating_Missing_Ind = "";

                            if ((itemBAP.Rating_Ori_Good_Ind ?? "") == "")
                            {
                                Match_Rating_Ori_Good_Ind = "Y";
                            }
                            else
                            {
                                if (tempRating_Ori_Good_IND.IsNullOrWhiteSpace() == true)
                                {
                                    tempBasic_Pass = "N";
                                }
                                else if (itemBAP.Rating_Ori_Good_Ind == tempRating_Ori_Good_IND)
                                {
                                    Match_Rating_Ori_Good_Ind = "Y";
                                }
                            }

                            if (itemBAP.Rating_Notch == null)
                            {
                                Match_Rating_Notch = "Y";
                            }
                            else
                            {
                                if (tempCurr_Ori_Rating_Diff.IsNullOrWhiteSpace() == true)
                                {
                                    tempBasic_Pass = "N";
                                }
                                else
                                {
                                    if (CompareRange(
                                                        itemBAP.Including_Ind,
                                                        itemBAP.Apply_Range,
                                                        int.Parse(tempCurr_Ori_Rating_Diff),
                                                        TypeTransfer.intNToInt(itemBAP.Rating_Notch)
                                                    ) == "Y")
                                    {
                                        Match_Rating_Notch = "Y";
                                    }
                                }
                            }

                            if ((itemBAP.Rating_Curr_Good_Ind ?? "") == "")
                            {
                                Match_Rating_Curr_Good_Ind = "Y";
                            }
                            else
                            {
                                if (tempRating_Curr_Good_Ind.IsNullOrWhiteSpace() == true)
                                {
                                    tempBasic_Pass = "N";
                                }
                                else if (itemBAP.Rating_Curr_Good_Ind == tempRating_Curr_Good_Ind)
                                {
                                    Match_Rating_Curr_Good_Ind = "Y";
                                }
                            }

                            if ((itemBAP.Ori_Rating_Missing_Ind ?? "") == "")
                            {
                                Match_Ori_Rating_Missing_Ind = "Y";
                            }
                            else
                            {
                                if (itemBAP.Ori_Rating_Missing_Ind.ToString() == "N")
                                {
                                    if (tempOriginal_External_Rating.IsNullOrWhiteSpace() == false)
                                    {
                                        Match_Ori_Rating_Missing_Ind = "Y";
                                    }
                                }
                                else if (itemBAP.Ori_Rating_Missing_Ind.ToString() == "Y")
                                {
                                    if (tempOriginal_External_Rating.IsNullOrWhiteSpace() == true)
                                    {
                                        Match_Ori_Rating_Missing_Ind = "Y";
                                    }
                                }
                            }

                            if (Match_Rating_Ori_Good_Ind == "Y" && Match_Rating_Notch == "Y" && Match_Rating_Curr_Good_Ind == "Y" && Match_Ori_Rating_Missing_Ind == "Y") 
                            {
                                tempBasic_Pass = itemBAP.Basic_Pass;
                                tempMap_Rule_Id_D69 = itemBAP.Rule_ID.ToString();

                                break;
                            }
                        }

                        if (tempCost_Value.IsNullOrWhiteSpace() == false && tempMarket_Value_Ori.IsNullOrWhiteSpace() == false)
                        {
                            if (double.Parse(tempCost_Value) != 0)
                            {
                                tempValue_Change_Ratio = ((double.Parse(tempMarket_Value_Ori) - double.Parse(tempCost_Value)) / double.Parse(tempCost_Value)).ToString();
                            }
                        }

                        if (tempValue_Change_Ratio.IsNullOrWhiteSpace() == false)
                        {
                            if (double.Parse(tempValue_Change_Ratio) < -0.3)
                            {
                                tempValue_Change_Ratio_Pass = "Y";
                            }
                            else
                            {
                                tempValue_Change_Ratio_Pass = "N";
                            }
                        }

                        //var oneA96 = ListA96.Where(x =>
                        //x.Reference_Nbr == tempReference_Nbr
                        //).FirstOrDefault();
                        var oneA96 = ListA96.Where(x =>
                                     x.Bond_Number == tempBond_Number &&
                                     x.Lots == tempLots &&
                                     x.Portfolio_Name == tempPortfolio_Name                    
                                     ).FirstOrDefault();
                        if (oneA96 != null)
                        {
                            tempChg_In_Spread = oneA96.Chg_In_Spread.ToString();
                        }

                        if (tempChg_In_Spread.IsNullOrWhiteSpace() == false)
                        {
                            tempSpread_Change_Over = getD62Spread_Change_Over(tempChg_In_Spread);
                        }

                        Bond_Basic_Assessment bba = bbaList.Where(x => x.Bond_Number == tempBond_Number
                                                                    && x.Lots == tempLots
                                                                    && x.Portfolio_Name == tempPortfolio_Name
                                                                    && x.Report_Date == lastReportDate)
                                                           .OrderByDescending(x => x.Version)
                                                           .FirstOrDefault(); 
                        if (bba != null)
                        {
                            tempAccumulation_Loss_last_Month = bba.Accumulation_Loss_This_Month.ToString();
                            tempChg_In_Spread_last_Month = bba.Chg_In_Spread_This_Month.ToString();
                        }

                        tempAccumulation_Loss_last_Month = tempAccumulation_Loss_last_Month ?? "";
                        tempChg_In_Spread_last_Month = tempChg_In_Spread_last_Month ?? "";

                        if (tempValue_Change_Ratio_Pass == "N")
                        {
                            tempAccumulation_Loss_This_Month = "0";
                        }
                        else
                        {
                            if (tempAccumulation_Loss_last_Month.IsNullOrWhiteSpace() == false)
                            {
                                tempAccumulation_Loss_This_Month = (int.Parse(tempAccumulation_Loss_last_Month) + 1).ToString();
                            }
                        }

                        tempChg_In_Spread_This_Month = getD62Chg_In_Spread_This_Month(tempSpread_Change_Over, tempChg_In_Spread_last_Month);

                        Bond_Basic_Assessment bba2 = bbaList.Where(x => x.Bond_Number == tempBond_Number
                                              && x.Lots == tempLots
                                              && x.Portfolio_Name == tempPortfolio_Name
                                              && x.Report_Date == lastTwoReportDate)
                                     .OrderByDescending(x => x.Version)
                                     .FirstOrDefault();
                        if (bba2 != null)
                        {
                            tempLastTwoBasicPass = bba2.Basic_Pass;
                            tempLastTwoCurrentExternalRating = bba2.Current_External_Rating.ToString();
                        }

                        //bba = bbaList.Where(x => x.Bond_Number == tempBond_Number
                        //                      && x.Lots == tempLots
                        //                      && x.Portfolio_Name == tempPortfolio_Name
                        //                      && x.Report_Date == lastReportDate)
                        //             .OrderByDescending(x => x.Version)
                        //             .FirstOrDefault();
                        if (bba != null)
                        {
                            tempLastBasicPass = bba.Basic_Pass;
                            tempLastCurrentExternalRating = bba.Current_External_Rating.ToString();
                        }


                        tempWatch_IND = "N";
                        var queryD70Result = getD62Watch_IND(ListWatch,
                                                             tempLastTwoBasicPass,
                                                             tempLastBasicPass,
                                                             tempBasic_Pass,
                                                             tempLastTwoCurrentExternalRating,
                                                             tempLastCurrentExternalRating,
                                                             tempCurrent_External_Rating,
                                                             tempAccumulation_Loss_This_Month,
                                                             tempChg_In_Spread_This_Month);
                        tempWatch_IND = queryD70Result.Item1;
                        tempMap_Rule_Id_D70 = queryD70Result.Item2;

                        tempWarning_IND = "N";
                        var queryD71Result = getD62Warning_IND(ListWarning,
                                                               tempLastTwoBasicPass,
                                                               tempLastBasicPass,
                                                               tempBasic_Pass,
                                                               tempLastTwoCurrentExternalRating,
                                                               tempLastCurrentExternalRating,
                                                               tempCurrent_External_Rating,
                                                               tempAccumulation_Loss_This_Month,
                                                               tempChg_In_Spread_This_Month);
                        tempWarning_IND = queryD71Result.Item1;
                        tempMap_Rule_Id_D71 = queryD71Result.Item2;

                        sb.Append(string.Format(@"INSERT INTO Bond_Basic_Assessment(Reference_Nbr,
                                                                                    Version,
                                                                                    Report_Date,
                                                                                    Processing_Date,
                                                                                    Bond_Type,
                                                                                    Assessment_Sub_Kind,
                                                                                    Lien_position,
                                                                                    Bond_Number,
                                                                                    Lots,
                                                                                    Portfolio,
                                                                                    Portfolio_Name,
                                                                                    Origination_Date,
                                                                                    Original_External_Rating,
                                                                                    Current_External_Rating,
                                                                                    Rating_Ori_Good_IND,
                                                                                    Rating_Curr_Good_Ind,
                                                                                    Curr_Ori_Rating_Diff,
                                                                                    Basic_Pass,
                                                                                    Map_Rule_Id_D69,
                                                                                    Cost_Value,
                                                                                    Market_Value_Ori,
                                                                                    Value_Change_Ratio,
                                                                                    Value_Change_Ratio_Pass,
                                                                                    Accumulation_Loss_last_Month,
                                                                                    Accumulation_Loss_This_Month,
                                                                                    Watch_IND,
                                                                                    Map_Rule_Id_D70,
                                                                                    Warning_IND,
                                                                                    Map_Rule_Id_D71,
                                                                                    Chg_In_Spread,
                                                                                    Spread_Change_Over,
                                                                                    Chg_In_Spread_last_Month,
                                                                                    Chg_In_Spread_This_Month)
                                                                         VALUES('{0}',{1},'{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}',{11},{12},{13},'{14}',
                                                                                '{15}',{16},'{17}',{18},{19},{20},{21},'{22}',{23},{24},'{25}','{26}','{27}','{28}',
                                                                                 {29},'{30}',{31},{32});",
                                                                                tempReference_Nbr,
                                                                                tempVersion,
                                                                                tempReport_Date,
                                                                                tempProcessing_Date,
                                                                                tempBond_Type,
                                                                                tempAssessment_Sub_Kind,
                                                                                tempLien_position,
                                                                                tempBond_Number,
                                                                                tempLots,
                                                                                tempPortfolio,
                                                                                tempPortfolio_Name,
                                                                                tempOrigination_Date.IsNullOrWhiteSpace() == false ? "'" + tempOrigination_Date + "'" : "NULL",
                                                                                tempOriginal_External_Rating.IsNullOrWhiteSpace() == false ? tempOriginal_External_Rating : "NULL",
                                                                                tempCurrent_External_Rating.IsNullOrWhiteSpace() == false ? tempCurrent_External_Rating : "NULL",
                                                                                tempRating_Ori_Good_IND,
                                                                                tempRating_Curr_Good_Ind,
                                                                                tempCurr_Ori_Rating_Diff.IsNullOrWhiteSpace() == false ? tempCurr_Ori_Rating_Diff : "NULL",
                                                                                tempBasic_Pass,
                                                                                tempMap_Rule_Id_D69.IsNullOrWhiteSpace() == false ? tempMap_Rule_Id_D69 : "NULL",
                                                                                tempCost_Value.IsNullOrWhiteSpace() == false ? tempCost_Value : "NULL",
                                                                                tempMarket_Value_Ori.IsNullOrWhiteSpace() == false ? tempMarket_Value_Ori : "NULL",
                                                                                tempValue_Change_Ratio.IsNullOrWhiteSpace() == false ? tempValue_Change_Ratio : "NULL",
                                                                                tempValue_Change_Ratio_Pass,
                                                                                tempAccumulation_Loss_last_Month.IsNullOrWhiteSpace() == false ? tempAccumulation_Loss_last_Month : "NULL",
                                                                                tempAccumulation_Loss_This_Month.IsNullOrWhiteSpace() == false ? tempAccumulation_Loss_This_Month : "NULL",
                                                                                tempWatch_IND,
                                                                                tempMap_Rule_Id_D70,
                                                                                tempWarning_IND,
                                                                                tempMap_Rule_Id_D71,
                                                                                tempChg_In_Spread.IsNullOrWhiteSpace() == false ? tempChg_In_Spread : "NULL",
                                                                                tempSpread_Change_Over,
                                                                                tempChg_In_Spread_last_Month.IsNullOrWhiteSpace() == false ? tempChg_In_Spread_last_Month : "NULL",
                                                                                tempChg_In_Spread_This_Month.IsNullOrWhiteSpace() == false ? tempChg_In_Spread_This_Month : "NULL"));
                    }

                    db.Database.ExecuteSqlCommand(sb.ToString());

                    result.RETURN_FLAG = true;

                    result.DESCRIPTION = getD62Message(reportDate, reportDate);
                }
                catch (Exception ex)
                {
                    result.RETURN_FLAG = false;
                    result.DESCRIPTION = Message_Type
                                         .save_Fail.GetDescription("D62",ex.Message);
                }
            }

            return result;
        }

        #endregion

        #region checkD70D71
        public MSGReturnModel checkD70D71()
        {
            MSGReturnModel result = new MSGReturnModel();

            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                try
                {
                    var D70Data = db.Watching_List_Parm.AsNoTracking().ToList();
                    if (!D70Data.Any())
                    {
                        result.RETURN_FLAG = false;
                        result.DESCRIPTION = "D70：觀察名單參數檔 無資料";
                        return result;
                    }
                    //else if (D70Data.Where(x => x.Status != "2").Any())
                    //{
                    //    result.RETURN_FLAG = false;
                    //    result.DESCRIPTION = "D70：觀察名單參數檔 規則未複核完成";
                    //    return result;
                    //}
                    else if (D70Data.Where(x => x.IsActive == "Y").Any() == false)
                    {
                        result.RETURN_FLAG = false;
                        result.DESCRIPTION = "D70：觀察名單參數檔 沒有生效的規則";
                        return result;
                    }

                    var D71Data = db.Warning_List_Parm.AsNoTracking().ToList();
                    if (!D71Data.Any())
                    {
                        result.RETURN_FLAG = false;
                        result.DESCRIPTION = "D71：預警名單參數檔 無資料";
                        return result;
                    }
                    //else if (D71Data.Where(x => x.Status != "2").Any())
                    //{
                    //    result.RETURN_FLAG = false;
                    //    result.DESCRIPTION = "D71：預警名單參數檔 規則未複核完成";
                    //    return result;
                    //}
                    else if (D71Data.Where(x => x.IsActive == "Y").Any() == false)
                    {
                        result.RETURN_FLAG = false;
                        result.DESCRIPTION = "D71：預警名單參數檔 沒有生效的規則";
                        return result;
                    }

                    result.RETURN_FLAG = true;
                    result.DESCRIPTION = "";
                }
                catch (DbUpdateException ex)
                {
                    result.RETURN_FLAG = false;
                    result.DESCRIPTION = ex.Message;
                }
            }

            return result;
        }
        #endregion

        #region getD62Spread_Change_Over
        public string getD62Spread_Change_Over(string Chg_In_Spread)
        {
            string Spread_Change_Over = "";

            Chg_In_Spread = Chg_In_Spread ?? "";

            if (Chg_In_Spread.IsNullOrWhiteSpace() == false)
            {
                if (double.Parse(Chg_In_Spread) >= 300)
                {
                    Spread_Change_Over = "Y";
                }
                else
                {
                    Spread_Change_Over = "N";
                }
            }

            return Spread_Change_Over;
        }
        #endregion

        #region getD62Chg_In_Spread_This_Month
        public string getD62Chg_In_Spread_This_Month(string Spread_Change_Over, string Chg_In_Spread_last_Month)
        {
            string Chg_In_Spread_This_Month = "";

            Spread_Change_Over = Spread_Change_Over ?? "";
            Chg_In_Spread_last_Month = Chg_In_Spread_last_Month ?? "";

            if (Spread_Change_Over.IsNullOrWhiteSpace() == false)
            {
                if (Spread_Change_Over == "N")
                {
                    Chg_In_Spread_This_Month = "0";
                }
                else
                {
                    if (Chg_In_Spread_last_Month.IsNullOrWhiteSpace() == false)
                    {
                        Chg_In_Spread_This_Month = (int.Parse(Chg_In_Spread_last_Month) + 1).ToString();
                    }
                    else
                    {
                        Chg_In_Spread_This_Month = "1";
                    }
                }
            }
            else
            {
                if (Chg_In_Spread_last_Month.IsNullOrWhiteSpace() == false)
                {
                    Chg_In_Spread_This_Month = (int.Parse(Chg_In_Spread_last_Month) + 1).ToString();
                }
            }

            return Chg_In_Spread_This_Month;
        }
        #endregion

        #region getD62Watch_IND
        public Tuple<string, string> getD62Watch_IND(List<Watching_List_Parm> D70Data,
                                                     string LastTwoBasic_Pass,
                                                     string LastBasic_Pass,
                                                     string ThisBasic_Pass,
                                                     string LastTwoCurrent_External_Rating,
                                                     string LastCurrent_External_Rating,
                                                     string ThisCurrent_External_Rating,
                                                     string Accumulation_Loss_This_Month,
                                                     string Chg_In_Spread_This_Month)
        {
            LastTwoBasic_Pass = LastTwoBasic_Pass ?? "";
            LastBasic_Pass = LastBasic_Pass ?? "";
            ThisBasic_Pass = ThisBasic_Pass ?? "";
            LastTwoCurrent_External_Rating = LastTwoCurrent_External_Rating ?? "";
            LastCurrent_External_Rating = LastCurrent_External_Rating ?? "";
            ThisCurrent_External_Rating = ThisCurrent_External_Rating ?? "";
            Accumulation_Loss_This_Month = Accumulation_Loss_This_Month ?? "";
            Chg_In_Spread_This_Month = Chg_In_Spread_This_Month ?? "";

            string compareBasic_Pass = "";
            string compareCurrent_External_Rating = "";
            string Watch_IND = "N";
            string Map_Rule_Id_D70 = "";

            foreach (Watching_List_Parm itemWLP in D70Data)
            {
                string Match_Basic_Pass = "";
                string Match_Rating_Threshold_Map_Grade_Adjust = "";
                string Match_Rating_Map_Grade_Adjust = "";
                string Match_Value_Change_Months = "";
                string Match_Spread_Change_Months = "";

                switch (itemWLP.Basic_Pass_Check_Point)
                {
                    case "-2":
                        compareBasic_Pass = LastTwoBasic_Pass;
                        break;
                    case "-1":
                        compareBasic_Pass = LastBasic_Pass;
                        break;
                    case "0":
                        compareBasic_Pass = ThisBasic_Pass;
                        break;
                    default:
                        break;
                }

                switch (itemWLP.Rating_Check_Point)
                {
                    case "-2":
                        compareCurrent_External_Rating = LastTwoCurrent_External_Rating;
                        break;
                    case "-1":
                        compareCurrent_External_Rating = LastCurrent_External_Rating;
                        break;
                    case "0":
                        compareCurrent_External_Rating = ThisCurrent_External_Rating;
                        break;
                    default:
                        break;
                }

                if (itemWLP.Basic_Pass == "-999")
                {
                    if (compareBasic_Pass == "")
                    {
                        Match_Basic_Pass = "Y";
                    }
                }
                else
                {
                    if (compareBasic_Pass == itemWLP.Basic_Pass)
                    {
                        Match_Basic_Pass = "Y";
                    }
                }

                if (itemWLP.Rating_Threshold == "-999")
                {
                    if (compareCurrent_External_Rating == "")
                    {
                        Match_Rating_Threshold_Map_Grade_Adjust = "Y";
                    }
                }
                else if (itemWLP.Rating_Threshold_Map_Grade_Adjust == null)
                {
                    Match_Rating_Threshold_Map_Grade_Adjust = "Y";
                }
                else
                {
                    if (compareCurrent_External_Rating.IsNullOrWhiteSpace() == true)
                    {
                        //Watch_IND = "Y";
                    }
                    else
                    {
                        if (CompareRange(itemWLP.Including_Ind_0,
                                         itemWLP.Apply_Range_0,
                                         TypeTransfer.intNToInt(itemWLP.Rating_Threshold_Map_Grade_Adjust),
                                         int.Parse(compareCurrent_External_Rating)
                                        ) == "Y")
                        {
                            Match_Rating_Threshold_Map_Grade_Adjust = "Y";
                        }
                    }
                }

                if (itemWLP.Rating_from == "-999" || itemWLP.Rating_To == "-999")
                {
                    if (compareCurrent_External_Rating == "")
                    {
                        Match_Rating_Map_Grade_Adjust = "Y";
                    }
                }
                else if (itemWLP.Rating_from_Map_Grade_Adjust == null || itemWLP.Rating_To_Map_Grade_Adjust == null)
                {
                    Match_Rating_Map_Grade_Adjust = "Y";
                }
                else
                {
                    if (compareCurrent_External_Rating.IsNullOrWhiteSpace() == true)
                    {
                        //Watch_IND = "Y";
                    }
                    else
                    {
                        if (int.Parse(compareCurrent_External_Rating) >= itemWLP.Rating_from_Map_Grade_Adjust
                            && int.Parse(compareCurrent_External_Rating) <= itemWLP.Rating_To_Map_Grade_Adjust)
                        {
                            Match_Rating_Map_Grade_Adjust = "Y";
                        }
                    }
                }

                if (itemWLP.Value_Change_Months == -999)
                {
                    if (Accumulation_Loss_This_Month == "")
                    {
                        Match_Value_Change_Months = "Y";
                    }
                }
                else if (itemWLP.Value_Change_Months == null)
                {
                    Match_Value_Change_Months = "Y";
                }
                else
                {
                    if (Accumulation_Loss_This_Month.IsNullOrWhiteSpace() == true)
                    {
                        //Watch_IND = "Y";
                    }
                    else
                    {
                        if (CompareRange(itemWLP.Including_Ind_1,
                                         itemWLP.Apply_Range_1,
                                         int.Parse(Accumulation_Loss_This_Month),
                                         TypeTransfer.intNToInt(itemWLP.Value_Change_Months)
                                         ) == "Y")
                        {
                            Match_Value_Change_Months = "Y";
                        }
                    }
                }

                if (itemWLP.Spread_Change_Months == -999)
                {
                    if (Chg_In_Spread_This_Month == "")
                    {
                        Match_Spread_Change_Months = "Y";
                    }
                }
                else if (itemWLP.Spread_Change_Months == null)
                {
                    Match_Spread_Change_Months = "Y";
                }
                else
                {
                    if (Chg_In_Spread_This_Month.IsNullOrWhiteSpace() == true)
                    {
                        //Watch_IND = "Y";
                    }
                    else
                    {
                        if (CompareRange(itemWLP.Including_Ind_2,
                                         itemWLP.Apply_Range_2,
                                         int.Parse(Chg_In_Spread_This_Month),
                                         TypeTransfer.intNToInt(itemWLP.Spread_Change_Months)
                                         ) == "Y")
                        {
                            Match_Spread_Change_Months = "Y";
                        }
                    }
                }

                if (Match_Basic_Pass == "Y"
                    && Match_Rating_Threshold_Map_Grade_Adjust == "Y" && Match_Rating_Map_Grade_Adjust == "Y"
                    && Match_Value_Change_Months == "Y" && Match_Spread_Change_Months == "Y")
                {
                    Watch_IND = "Y";
                    Map_Rule_Id_D70 += itemWLP.Rule_ID.ToString() + ",";
                }
            }

            if (Map_Rule_Id_D70 != "")
            {
                Map_Rule_Id_D70 = Map_Rule_Id_D70.Substring(0, Map_Rule_Id_D70.Length - 1);
            }

            return new Tuple<string, string>(Watch_IND, Map_Rule_Id_D70);
        }
        #endregion

        #region getD62Warning_IND
        public Tuple<string, string> getD62Warning_IND(List<Warning_List_Parm> D71Data,
                                                       string LastTwoBasic_Pass,
                                                       string LastBasic_Pass,
                                                       string ThisBasic_Pass,
                                                       string LastTwoCurrent_External_Rating,
                                                       string LastCurrent_External_Rating,
                                                       string ThisCurrent_External_Rating,
                                                       string Accumulation_Loss_This_Month,
                                                       string Chg_In_Spread_This_Month)
        {
            LastTwoBasic_Pass = LastTwoBasic_Pass ?? "";
            LastBasic_Pass = LastBasic_Pass ?? "";
            ThisBasic_Pass = ThisBasic_Pass ?? "";
            LastTwoCurrent_External_Rating = LastTwoCurrent_External_Rating ?? "";
            LastCurrent_External_Rating = LastCurrent_External_Rating ?? "";
            ThisCurrent_External_Rating = ThisCurrent_External_Rating ?? "";
            Accumulation_Loss_This_Month = Accumulation_Loss_This_Month ?? "";
            Chg_In_Spread_This_Month = Chg_In_Spread_This_Month ?? "";

            string compareBasic_Pass = "";
            string compareCurrent_External_Rating = "";
            string Warning_IND = "N";
            string Map_Rule_Id_D71 = "";

            foreach (Warning_List_Parm itemWLP in D71Data)
            {
                string Match_Basic_Pass = "";
                string Match_Rating_Threshold_Map_Grade_Adjust = "";
                string Match_Rating_Map_Grade_Adjust = "";
                string Match_Value_Change_Months = "";
                string Match_Spread_Change_Months = "";

                switch (itemWLP.Basic_Pass_Check_Point)
                {
                    case "-2":
                        compareBasic_Pass = LastTwoBasic_Pass;
                        break;
                    case "-1":
                        compareBasic_Pass = LastBasic_Pass;
                        break;
                    case "0":
                        compareBasic_Pass = ThisBasic_Pass;
                        break;
                    default:
                        break;
                }

                switch (itemWLP.Rating_Check_Point)
                {
                    case "-2":
                        compareCurrent_External_Rating = LastTwoCurrent_External_Rating;
                        break;
                    case "-1":
                        compareCurrent_External_Rating = LastCurrent_External_Rating;
                        break;
                    case "0":
                        compareCurrent_External_Rating = ThisCurrent_External_Rating;
                        break;
                    default:
                        break;
                }

                if (itemWLP.Basic_Pass == "-999")
                {
                    if (compareBasic_Pass == "")
                    {
                        Match_Basic_Pass = "Y";
                    }
                }
                else
                {
                    if (compareBasic_Pass == itemWLP.Basic_Pass)
                    {
                        Match_Basic_Pass = "Y";
                    }
                }

                if (itemWLP.Rating_Threshold == "-999")
                {
                    if (compareCurrent_External_Rating == "")
                    {
                        Match_Rating_Threshold_Map_Grade_Adjust = "Y";
                    }
                }
                else if (itemWLP.Rating_Threshold_Map_Grade_Adjust == null)
                {
                    Match_Rating_Threshold_Map_Grade_Adjust = "Y";
                }
                else
                {
                    if (compareCurrent_External_Rating.IsNullOrWhiteSpace() == true)
                    {
                        //Warning_IND = "Y";
                    }
                    else
                    {
                        if (CompareRange(itemWLP.Including_Ind_0,
                                         itemWLP.Apply_Range_0,
                                         TypeTransfer.intNToInt(itemWLP.Rating_Threshold_Map_Grade_Adjust),
                                         int.Parse(compareCurrent_External_Rating)
                                        ) == "Y")
                        {
                            Match_Rating_Threshold_Map_Grade_Adjust = "Y";
                        }
                    }
                }

                if (itemWLP.Rating_from == "-999" || itemWLP.Rating_To == "-999")
                {
                    if (compareCurrent_External_Rating == "")
                    {
                        Match_Rating_Map_Grade_Adjust = "Y";
                    }
                }
                else if (itemWLP.Rating_from_Map_Grade_Adjust == null || itemWLP.Rating_To_Map_Grade_Adjust == null)
                {
                    Match_Rating_Map_Grade_Adjust = "Y";
                }
                else
                {
                    if (compareCurrent_External_Rating.IsNullOrWhiteSpace() == true)
                    {
                        //Warning_IND = "Y";
                    }
                    else
                    {
                        if (int.Parse(compareCurrent_External_Rating) >= itemWLP.Rating_from_Map_Grade_Adjust
                            && int.Parse(compareCurrent_External_Rating) <= itemWLP.Rating_To_Map_Grade_Adjust)
                        {
                            Match_Rating_Map_Grade_Adjust = "Y";
                        }
                    }
                }

                if (itemWLP.Value_Change_Months == -999)
                {
                    if (Accumulation_Loss_This_Month == "")
                    {
                        Match_Value_Change_Months = "Y";
                    }
                }
                else if (itemWLP.Value_Change_Months == null)
                {
                    Match_Value_Change_Months = "Y";
                }
                else
                {
                    if (Accumulation_Loss_This_Month.IsNullOrWhiteSpace() == true)
                    {
                        //Warning_IND = "Y";
                    }
                    else
                    {
                        if (CompareRange(itemWLP.Including_Ind_1,
                                         itemWLP.Apply_Range_1,
                                         int.Parse(Accumulation_Loss_This_Month),
                                         TypeTransfer.intNToInt(itemWLP.Value_Change_Months)
                                         ) == "Y")
                        {
                            Match_Value_Change_Months = "Y";
                        }
                    }
                }

                if (itemWLP.Spread_Change_Months == -999)
                {
                    if (Chg_In_Spread_This_Month == "")
                    {
                        Match_Spread_Change_Months = "Y";
                    }
                }
                else if (itemWLP.Spread_Change_Months == null)
                {
                    Match_Spread_Change_Months = "Y";
                }
                else
                {
                    if (Chg_In_Spread_This_Month.IsNullOrWhiteSpace() == true)
                    {
                        //Warning_IND = "Y";
                    }
                    else
                    {
                        if (CompareRange(itemWLP.Including_Ind_2,
                                         itemWLP.Apply_Range_2,
                                         int.Parse(Chg_In_Spread_This_Month),
                                         TypeTransfer.intNToInt(itemWLP.Spread_Change_Months)
                                         ) == "Y")
                        {
                            Match_Spread_Change_Months = "Y";
                        }
                    }
                }

                if (Match_Basic_Pass == "Y"
                    && Match_Rating_Threshold_Map_Grade_Adjust == "Y" && Match_Rating_Map_Grade_Adjust == "Y"
                    && Match_Value_Change_Months == "Y" && Match_Spread_Change_Months == "Y")
                {
                    Warning_IND = "Y";
                    Map_Rule_Id_D71 += itemWLP.Rule_ID.ToString() + ",";
                }
            }

            if (Map_Rule_Id_D71 != "")
            {
                Map_Rule_Id_D71 = Map_Rule_Id_D71.Substring(0, Map_Rule_Id_D71.Length - 1);
            }

            return new Tuple<string, string>(Warning_IND, Map_Rule_Id_D71);
        }
        #endregion

        #region getD62ByChg_In_Spread
        public Tuple<string, D62ViewModel> getD62ByChg_In_Spread(D62ViewModel dataModel)
        {
            D62ViewModel D62 = new D62ViewModel();

            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                if (dataModel.Chg_In_Spread.IsNullOrWhiteSpace() == false)
                {
                    D62.Spread_Change_Over = getD62Spread_Change_Over(dataModel.Chg_In_Spread);
                }

                if (D62.Spread_Change_Over.IsNullOrWhiteSpace() == false)
                {
                    D62.Chg_In_Spread_This_Month = getD62Chg_In_Spread_This_Month(D62.Spread_Change_Over, dataModel.Chg_In_Spread_last_Month);
                }

                MSGReturnModel result = checkD70D71();
                if (result.RETURN_FLAG == false)
                {
                    return new Tuple<string, D62ViewModel>(result.DESCRIPTION, new D62ViewModel());
                }

                DateTime dateReportDate = DateTime.Parse(dataModel.Report_Date);
                DateTime lastReportDate = dateReportDate.AddDays(-dateReportDate.Day);
                DateTime lastTwoReportDate = dateReportDate.AddMonths(-1).AddDays(-dateReportDate.AddMonths(-1).Day);

                Bond_Basic_Assessment oneD62 = db.Bond_Basic_Assessment.Where(x => x.Reference_Nbr == dataModel.Reference_Nbr).FirstOrDefault();

                string LastTwoBasicPass = "";
                string LastBasicPass = "";
                string ThisBasicPass = oneD62.Basic_Pass;
                string LastTwoCurrentExternalRating = "";
                string LastCurrentExternalRating = "";
                string ThisCurrentExternalRating = oneD62.Current_External_Rating.ToString();

                List<Bond_Basic_Assessment> bbaList = db.Bond_Basic_Assessment.AsNoTracking()
                                                        .Where(x => x.Bond_Number == dataModel.Bond_Number
                                                                 && x.Lots == dataModel.Lots
                                                                 && x.Portfolio_Name == dataModel.Portfolio_Name)
                                                        .ToList();

                var bba = bbaList.Where(x => x.Report_Date == lastTwoReportDate)
                                 .OrderByDescending(x => x.Version)
                                 .FirstOrDefault();
                if (bba != null)
                {
                    LastTwoBasicPass = bba.Basic_Pass;
                    LastTwoCurrentExternalRating = bba.Current_External_Rating.ToString();
                }

                bba = bbaList.Where(x => x.Report_Date == lastReportDate)
                             .OrderByDescending(x => x.Version)
                             .FirstOrDefault();
                if (bba != null)
                {
                    LastBasicPass = bba.Basic_Pass;
                    LastCurrentExternalRating = bba.Current_External_Rating.ToString();
                }

                D62.Watch_IND = "N";
                var D70Data = db.Watching_List_Parm.Where(x => x.IsActive == "Y").ToList();
                var queryD70Result = getD62Watch_IND(D70Data,
                                                     LastTwoBasicPass,
                                                     LastBasicPass,
                                                     ThisBasicPass,
                                                     LastTwoCurrentExternalRating,
                                                     LastCurrentExternalRating,
                                                     ThisCurrentExternalRating,
                                                     dataModel.Accumulation_Loss_This_Month,
                                                     D62.Chg_In_Spread_This_Month);
                D62.Watch_IND = queryD70Result.Item1;
                D62.Map_Rule_Id_D70 = queryD70Result.Item2;

                D62.Warning_IND = "N";
                var D71Data = db.Warning_List_Parm.Where(x => x.IsActive == "Y").ToList();
                var queryD71Result = getD62Warning_IND(D71Data,
                                                       LastTwoBasicPass,
                                                       LastBasicPass,
                                                       ThisBasicPass,
                                                       LastTwoCurrentExternalRating,
                                                       LastCurrentExternalRating,
                                                       ThisCurrentExternalRating,
                                                       dataModel.Accumulation_Loss_This_Month,
                                                       D62.Chg_In_Spread_This_Month);
                D62.Warning_IND = queryD71Result.Item1;
                D62.Map_Rule_Id_D71 = queryD71Result.Item2;
            }

            return new Tuple<string, D62ViewModel>("", D62);
        }
        #endregion

        #region modifyD62
        public MSGReturnModel modifyD62(D62ViewModel dataModel)
        {
            MSGReturnModel result = new MSGReturnModel();

            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                try
                {
                    Bond_Basic_Assessment oldData = db.Bond_Basic_Assessment.Where(x => x.Reference_Nbr == dataModel.Reference_Nbr)
                                                                            .FirstOrDefault();

                    oldData.Chg_In_Spread = TypeTransfer.stringToDoubleN(dataModel.Chg_In_Spread);
                    oldData.Spread_Change_Over = dataModel.Spread_Change_Over;
                    oldData.Chg_In_Spread_This_Month = TypeTransfer.stringToIntN(dataModel.Chg_In_Spread_This_Month);

                    if ((oldData.Watch_IND.ToString() != dataModel.Watch_IND.ToString()) 
                        || (oldData.Watch_IND_Override_Date.ToString().IsNullOrWhiteSpace() == false))
                    {
                        oldData.Watch_IND = dataModel.Watch_IND;
                        oldData.Map_Rule_Id_D70 = dataModel.Map_Rule_Id_D70;
                        oldData.Watch_IND_Override = dataModel.Watch_IND_Override;
                        oldData.Watch_IND_Override_Desc = dataModel.Watch_IND_Override_Desc;
                        oldData.Watch_IND_Override_Date = TypeTransfer.stringToDateTimeN(dataModel.Watch_IND_Override_Date);
                        oldData.Watch_IND_Override_User = dataModel.Watch_IND_Override_User;
                    }

                    if ((oldData.Warning_IND.ToString() != dataModel.Warning_IND.ToString())
                        || (oldData.Warning_IND_Override_Date.ToString().IsNullOrWhiteSpace() == false))
                    {
                        oldData.Warning_IND = dataModel.Warning_IND;
                        oldData.Map_Rule_Id_D71 = dataModel.Map_Rule_Id_D71;
                        oldData.Warning_IND_Override = dataModel.Warning_IND_Override;
                        oldData.Warning_IND_Override_DESC = dataModel.Warning_IND_Override_DESC;
                        oldData.Warning_IND_Override_Date = TypeTransfer.stringToDateTimeN(dataModel.Warning_IND_Override_Date);
                        oldData.Warning_IND_Override_User = dataModel.Warning_IND_Override_User;
                    }

                    db.SaveChanges();

                    result.RETURN_FLAG = true;
                }
                catch (DbUpdateException ex)
                {
                    result.RETURN_FLAG = false;
                    result.DESCRIPTION = ex.Message;
                }
            }

            return result;
        }
        #endregion

        #region Compare Range
        private string CompareRange(string Including_Ind, string Apply_Range, int number1, int number2)
        {
            string result = "N";

            if (Including_Ind == "Y")
            {
                if (Apply_Range == "1")
                {
                    if (number1 >= number2) { result = "Y"; }
                }
                else if (Apply_Range == "0")
                {
                    if (number1 <= number2) { result = "Y"; }
                }
            }
            else if (Including_Ind == "N")
            {
                if (Apply_Range == "1")
                {
                    if (number1 > number2) { result = "Y"; }
                }
                else if (Apply_Range == "0")
                {
                    if (number1 < number2) { result = "Y"; }
                }
            }

            return result;
        }

        #endregion

        //#region 下載 D64 Excel

        ///// <summary>
        ///// 下載 Excel
        ///// </summary>
        ///// <param name="path">下載位置</param>
        ///// <param name="dbDatas"></param>
        ///// <returns></returns>
        //public MSGReturnModel DownLoadD64Excel(string path, List<D64ViewModel> dbDatas)
        //{
        //    MSGReturnModel result = new MSGReturnModel();
        //    result.RETURN_FLAG = false;
        //    result.DESCRIPTION = Message_Type.download_Fail
        //                         .GetDescription("D64", Message_Type.not_Find_Any.GetDescription());

        //    if (dbDatas.Any())
        //    {
        //        DataTable datas = getD64ModelFromDb(dbDatas).Item1;

        //        result.DESCRIPTION = FileRelated.DataTableToExcel(datas, path, Excel_DownloadName.D64);
        //        result.RETURN_FLAG = string.IsNullOrWhiteSpace(result.DESCRIPTION);

        //        if (result.RETURN_FLAG)
        //        {
        //            result.DESCRIPTION = Message_Type.download_Success.GetDescription();
        //        }
        //    }

        //    return result;
        //}

        //#endregion

        #region Get Data

        #region D63評估執行成功 查詢
        /// <summary>
        /// D63評估執行成功 查詢
        /// </summary>
        /// <param name="reportDate"></param>
        /// <returns></returns>
        public List<D63ViewModel> TransferD63Data(string reportDate)
        {
            List<D63ViewModel> result = new List<D63ViewModel>();
            DateTime dt = DateTime.MinValue;
            int version = 0;
            if (!DateTime.TryParse(reportDate, out dt))
            {
                return result;
            }
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                var D62 = db.Bond_Basic_Assessment.AsNoTracking().Where(x => x.Report_Date == dt).ToList();
                var D63 = db.Bond_Quantitative_Resource.AsNoTracking()
                    .Where(x => x.Report_Date == dt);
                if (D62.Any())
                    version = D62.Max(x => x.Version);
                //version = D63.Max(x => x.Version);
                var Users = db.IFRS9_User.AsNoTracking().ToList();
                var _other = AssessmentSubKind_Type.Other.GetDescription();
                result.AddRange(
                    D63.Where(x => x.Version == version &&
                     x.Assessment_Result_Version == 1 &&
                     x.Assessment_Sub_Kind != _other)
                    .AsEnumerable()
                    .Select(x => DbToD63ViewModel(x,Users)));
                result.AddRange(
                    D63.Where(x => x.Version == version &&
                     x.Assessment_Result_Version == 2 &&
                     x.Assessment_Sub_Kind == _other)
                     .AsEnumerable()
                     .Select(x => DbToD63ViewModel(x,Users)));
            }
            result.ForEach(x =>
            {
                if (x.Assessment_Sub_Kind != AssessmentSubKind_Type.Other.GetDescription())
                    x.Status = Evaluation_Status_Type.NotAssess.GetDescription();
            });
            return result.OrderBy(x=>x.Reference_Nbr).ToList();
        }
        #endregion

        #region D65評估執行成功 查詢
        /// <summary>
        /// D65評估執行成功 查詢
        /// </summary>
        /// <param name="reportDate"></param>
        /// <returns></returns>
        public List<D65ViewModel> TransferD65Data(string reportDate)
        {
            List<D65ViewModel> result = new List<D65ViewModel>();
            DateTime dt = DateTime.MinValue;
            int version = 0;
            if (!DateTime.TryParse(reportDate, out dt))
            {
                return result;
            }
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                var D62 = db.Bond_Basic_Assessment.AsNoTracking().Where(x => x.Report_Date == dt).ToList();                
                var D65 = db.Bond_Qualitative_Assessment.AsNoTracking()
                    .Where(x => x.Report_Date == dt);
                if (D62.Any())
                    version = D62.Max(x => x.Version);
                //version = D65.Max(x => x.Version);
                var Users = db.IFRS9_User.AsNoTracking().ToList();
                result = D65.Where(x => x.Version == version).AsEnumerable()
                    .Select(x=> DbToD65ViewModel(x, Users)).ToList();
                result.ForEach(x =>
                {
                    x.Status = Evaluation_Status_Type.NotAssess.GetDescription();
                });
            }
            return result;
        }
        #endregion

        #region get D63 評估結果版本
        /// <summary>
        /// get D63 評估結果版本
        /// </summary>
        /// <param name="reportDate">報導日</param>
        /// <param name="boundNumber">債券編號</param>
        /// <param name="indexFlag">評估狀態</param>
        /// <param name="assessmentSubKind">評估次分類</param>
        /// <param name="Send_to_AuditorFlag">是否提交複核</param>
        /// <param name="referenceNbr">帳戶編號</param>
        /// <returns></returns>
        public List<D63ViewModel> getD63(
            DateTime reportDate,
            string boundNumber,
            Evaluation_Status_Type indexFlag,
            string assessmentSubKind,
            bool Send_to_AuditorFlag = false,
            string referenceNbr = ""
            )
        {
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                int version = 0;
                var D62Version = db.Bond_Basic_Assessment.AsNoTracking().Where(x => x.Report_Date == reportDate).Select(x=>x.Version).ToList();
                if (!D62Version.Any())
                {
                    return new List<D63ViewModel>();
                }
                version = D62Version.Max();
                var D63Data = db.Bond_Quantitative_Resource.AsNoTracking().Where(x =>
                            x.Report_Date == reportDate && x.Version == version)
                            .Where(x => x.Bond_Number == boundNumber , !boundNumber.IsNullOrWhiteSpace()) //債券編號                                                              
                            .Where(x => x.Assessment_Sub_Kind == assessmentSubKind, assessmentSubKind != "All")  // 評估次分類
                            .Where(x => x.Send_to_Auditor == "Y", Send_to_AuditorFlag)
                            .Where(x => x.Reference_Nbr == referenceNbr, !referenceNbr.IsNullOrWhiteSpace())
                            .ToList();
                var D64s = db.Bond_Quantitative_Result.AsNoTracking().Where(x =>
                            x.Report_Date == reportDate &&
                            x.Version == version).ToList();
                var D64Files = db.Bond_Quantitative_Result_File.AsNoTracking().Where(x=>x.Status == "Y").ToList();
                if (D63Data.Any())
                {
                    var Users = common.getAllUsers();
                    var results = D63Data.Select(x => DbToD63ViewModel(x, Users)).ToList();
                    List<D63ViewModel> result = new List<D63ViewModel>();
                    if (Send_to_AuditorFlag)
                        result = results.OrderByDescending(x => x.Assessment_Result_Version).ToList();
                    else
                    result = results.GroupBy(x => x.Reference_Nbr)
                            .Select(x => x.OrderByDescending(y => y.Auditor_Return != "Y")
                            .ThenByDescending(y => y.Send_to_Auditor == "Y")
                            .ThenByDescending(y => y.Assessment_Result_Version).FirstOrDefault()).ToList();
                    result.ForEach(x =>
                      {
                          var list = results.Where(y=>y.Reference_Nbr == x.Reference_Nbr).ToList();
                          var _status = getD63Status(list);
                          var rejectFlag =
                            (_status == Evaluation_Status_Type.ReviewCompleted ||
                             _status == Evaluation_Status_Type.Review) ? false :
                             (list.Count >= 2 && list.Any(z => z.Auditor_Return == "Y"));
                          list.ForEach(y =>
                          {
                              y.Status = y.Auditor_Return == "Y" ?
                              Evaluation_Status_Type.Reject.GetDescription() :
                              rejectFlag ? "複核被退回請重新提交" : _status.GetDescription();
                              y.Result_Version_Confirm_Flag = getResultVersionConfirmFlag(_status);
                          });
                      });
                    if ((indexFlag & Evaluation_Status_Type.NotAssess) == Evaluation_Status_Type.NotAssess)
                        result = result.Where(x => x.Status == Evaluation_Status_Type.NotAssess.GetDescription()).ToList();
                    if ((indexFlag & Evaluation_Status_Type.NotReview) == Evaluation_Status_Type.NotReview)
                        result = result.Where(x => x.Status == Evaluation_Status_Type.NotReview.GetDescription()).ToList();
                    if ((indexFlag & Evaluation_Status_Type.Reject) == Evaluation_Status_Type.Reject)
                    {
                        if (Send_to_AuditorFlag)
                            result = result.Where(x => x.Status == Evaluation_Status_Type.Reject.GetDescription()).ToList();
                        else
                            result = result.Where(x => x.Status == "複核被退回請重新提交").ToList();
                    }
                    if ((indexFlag & Evaluation_Status_Type.Review) == Evaluation_Status_Type.Review)
                        result = result.Where(x => x.Status == Evaluation_Status_Type.Review.GetDescription()).ToList();
                    if ((indexFlag & Evaluation_Status_Type.ReviewCompleted) == Evaluation_Status_Type.ReviewCompleted)
                        result = result.Where(x => x.Status == Evaluation_Status_Type.ReviewCompleted.GetDescription()).ToList();
                    result.ForEach(x =>
                    {
                        var _D64 = D64s.Where(y => y.Reference_Nbr == x.Reference_Nbr && y.Assessment_Result_Version.ToString() == x.Assessment_Result_Version);
                        var _count = 0;
                        foreach (var item in _D64)
                        {
                            _count += D64Files.Count(z => z.Check_Reference == item.Check_Reference);
                        }
                        x.FilesCount = _count.ToString();
                    });
                    return result;
                }
            }
            return new List<D63ViewModel>();
        }
        #endregion

        #region get D63 All Assessment_Result_Version 
        /// <summary>
        /// get D63 All Assessment_Result_Version 
        /// </summary>
        /// <param name="Reference_Nbr"></param>
        /// <param name="version"></param>
        /// <returns></returns>
        public List<D63ViewModel> getD63History(string Reference_Nbr)
        {
            var datas = new List<D63ViewModel>();
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                var history = db.Bond_Quantitative_Resource.AsNoTracking()
                    .Where(x => x.Reference_Nbr == Reference_Nbr)
                    .OrderBy(x => x.Assessment_Result_Version);
                if (history.Any())
                {
                    var Users = common.getAllUsers();
                    datas = history.AsEnumerable()
                        .OrderByDescending(x=>x.Assessment_Result_Version)
                        .Select(x => DbToD63ViewModel(x, Users)).ToList();
                    var _status = getD63Status(datas);
                    datas.ForEach(x =>
                    {
                        x.Status = 
                        (x.Assessment_Result_Version == "1") ? "第一版不複核" :
                        x.Auditor_Return == "Y" ?
                        Evaluation_Status_Type.Reject.GetDescription() : _status.GetDescription();
                        x.Result_Version_Confirm_Flag = getResultVersionConfirmFlag(_status);
                    });
                }
            }
            return datas;
        }
        #endregion

        #region get D64 Bond_Quantitative_Result
        /// <summary>
        /// get D64-量化評估結果檔
        /// </summary>
        /// <param name="referenceNbr"></param>
        /// <param name="Assessment_Result_Version"></param>
        /// <returns></returns>
        public List<D64ViewModel> getD64(string referenceNbr,int Assessment_Result_Version)
        {
            var datas = new List<D64ViewModel>();
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                var _Assessor = db.Bond_Quantitative_Resource.AsNoTracking()
                    .Where(x => x.Reference_Nbr == referenceNbr && x.Assessment_Result_Version == Assessment_Result_Version)
                    .DefaultIfEmpty().First().Assessor;
                var D64Files = db.Bond_Quantitative_Result_File.AsNoTracking().ToList();
                datas = db.Bond_Quantitative_Result.AsNoTracking()
                    .Where(x => x.Reference_Nbr == referenceNbr && x.Assessment_Result_Version == Assessment_Result_Version)
                    .AsEnumerable().Select(x => DbToD64ViewModel(x, _Assessor,D64Files)).ToList();
            }
            return datas;
        }
        #endregion

        #region get D65  評估結果版本
        /// <summary>
        /// get D65 評估結果版本
        /// </summary>
        /// <param name="reportDate">報導日</param>
        /// <param name="boundNumber">債券編號</param>
        /// <param name="referenceNbr">帳戶編號</param>
        /// <param name="indexFlag">評估狀態</param>
        /// <param name="assessmentSubKind">評估次分類</param>
        /// <param name="Send_to_AuditorFlag">是否提交複核</param>
        /// <returns></returns>
        public List<D65ViewModel> getD65(
            DateTime reportDate,
            string boundNumber,
            string referenceNbr,
            Evaluation_Status_Type indexFlag,
            string assessmentSubKind,
            bool Send_to_AuditorFlag = false
            )
        {
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                int version = 0;
                var D62Version = db.Bond_Basic_Assessment.AsNoTracking()
                    .Where(x => x.Report_Date == reportDate)
                    .Select(x => x.Version).ToList();
                if (!D62Version.Any())
                {
                    return new List<D65ViewModel>();
                }
                version = D62Version.Max();
                var D65Data = db.Bond_Qualitative_Assessment.AsNoTracking()
                    .Where(x => x.Report_Date == reportDate && x.Version == version)
                    .Where(x => x.Bond_Number == boundNumber, !boundNumber.IsNullOrWhiteSpace()) //債券編號                                                              
                    .Where(x => x.Assessment_Sub_Kind == assessmentSubKind, assessmentSubKind != "All")  // 評估次分類
                    .Where(x => x.Send_to_Auditor == "Y" && x.Extra_Case != "D", Send_to_AuditorFlag)
                    .Where(x => x.Reference_Nbr == referenceNbr, !referenceNbr.IsNullOrWhiteSpace())
                    .ToList();
                var D66s = db.Bond_Qualitative_Assessment_Result.AsNoTracking()
                    .Where(x => x.Report_Date == reportDate && x.Version == version).ToList();
                var D66Files = db.Bond_Qualitative_Assessment_Result_File.AsNoTracking().Where(x => x.Status == "Y").ToList();
                if (D65Data.Any())
                {
                    var Users = common.getAllUsers();
                    var results = D65Data.Select(x => DbToD65ViewModel(x, Users)).ToList();
                    List<D65ViewModel> result = new List<D65ViewModel>();
                    if (Send_to_AuditorFlag)
                        result = results.OrderByDescending(x => x.Assessment_Result_Version).ToList();
                    else
                        result = results.GroupBy(x => x.Reference_Nbr)
                                .Select(x => x.OrderByDescending(y => y.Auditor_Return != "Y")
                                .ThenByDescending(y => y.Send_to_Auditor == "Y")
                                .ThenByDescending(y => y.Assessment_Result_Version).FirstOrDefault()).ToList();
                    result.ForEach(x =>
                    {
                        var list = results.Where(y => y.Reference_Nbr == x.Reference_Nbr).ToList();
                        var _status = getD65Status(list);
                        var rejectFlag =
                        (_status == Evaluation_Status_Type.ReviewCompleted ||
                         _status == Evaluation_Status_Type.Review) ? false :
                         (list.Count >= 2 && list.Any(z => z.Auditor_Return == "Y"));
                        list.ForEach(y =>
                        {
                            y.Status = y.Auditor_Return == "Y" ?
                            Evaluation_Status_Type.Reject.GetDescription() :
                            rejectFlag ? "複核被退回請重新提交" : _status.GetDescription();
                            y.Result_Version_Confirm_Flag = getResultVersionConfirmFlag(_status);
                        });
                    });
                    if ((indexFlag & Evaluation_Status_Type.NotAssess) == Evaluation_Status_Type.NotAssess)
                        result = result.Where(x => x.Status == Evaluation_Status_Type.NotAssess.GetDescription()).ToList();
                    if ((indexFlag & Evaluation_Status_Type.NotReview) == Evaluation_Status_Type.NotReview)
                        result = result.Where(x => x.Status == Evaluation_Status_Type.NotReview.GetDescription()).ToList();
                    if ((indexFlag & Evaluation_Status_Type.Reject) == Evaluation_Status_Type.Reject)
                    {
                        if(Send_to_AuditorFlag)
                            result = result.Where(x => x.Status == Evaluation_Status_Type.Reject.GetDescription()).ToList();
                        else
                            result = result.Where(x => x.Status == "複核被退回請重新提交").ToList();
                    }
                    if ((indexFlag & Evaluation_Status_Type.Review) == Evaluation_Status_Type.Review)
                        result = result.Where(x => x.Status == Evaluation_Status_Type.Review.GetDescription()).ToList();
                    if ((indexFlag & Evaluation_Status_Type.ReviewCompleted) == Evaluation_Status_Type.ReviewCompleted)
                        result = result.Where(x => x.Status == Evaluation_Status_Type.ReviewCompleted.GetDescription()).ToList();
                    result.ForEach(x =>
                    {
                        var _D66 = D66s.Where(y => y.Reference_Nbr == x.Reference_Nbr && y.Assessment_Result_Version.ToString() == x.Assessment_Result_Version);
                        var _count = 0;
                        foreach (var item in _D66)
                        {
                            _count += D66Files.Count(z => z.Check_Reference == item.Check_Reference);
                        }
                        x.FilesCount = _count.ToString();
                    });
                    return result;
                }
            }
            return new List<D65ViewModel>();
        }
        #endregion

        #region get D65 All Assessment_Result_Version
        /// <summary>
        /// get D65 All Assessment_Result_Version
        /// </summary>
        /// <param name="Reference_Nbr"></param>
        /// <returns></returns>
        public List<D65ViewModel> getD65History(string Reference_Nbr)
        {
            var datas = new List<D65ViewModel>();
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                var history = db.Bond_Qualitative_Assessment.AsNoTracking()
                    .Where(x => x.Reference_Nbr == Reference_Nbr)
                    .OrderBy(x => x.Assessment_Result_Version);
                if (history.Any())
                {
                    var user = common.getAllUsers();

                    datas = history.AsEnumerable()
                        .OrderByDescending(x => x.Assessment_Result_Version)
                        .Select(x => DbToD65ViewModel(x, user)).ToList();
                    var _status = getD65Status(datas);
                    datas.ForEach(x =>
                    {
                        x.Status = (x.Assessment_Result_Version == "1") ? "第一版不複核" :
                        x.Auditor_Return == "Y" ?
                        Evaluation_Status_Type.Reject.GetDescription() : _status.GetDescription();
                        x.Result_Version_Confirm_Flag = getResultVersionConfirmFlag(_status);
                    });
                }
            }
            return datas;
        }
        #endregion

        #region get D66 Bond_Qualitative_Assessment_Result
        /// <summary>
        /// get D66-質化評估結果檔
        /// </summary>
        /// <param name="referenceNbr">帳戶編號</param>
        /// <param name="Assessment_Result_Version">評估結果版本</param>
        /// <returns></returns>
        public List<D66ViewModel> getD66(
            string referenceNbr,
            int Assessment_Result_Version
            )
        {
            var datas = new List<D66ViewModel>();
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                var _Assessor = db.Bond_Qualitative_Assessment.AsNoTracking()
                    .Where(x => x.Reference_Nbr == referenceNbr && x.Assessment_Result_Version == Assessment_Result_Version)
                    .DefaultIfEmpty().First().Assessor;
                var D66Files = db.Bond_Qualitative_Assessment_Result_File.AsNoTracking().ToList();
                datas = db.Bond_Qualitative_Assessment_Result.AsNoTracking()
                    .Where(x => x.Reference_Nbr == referenceNbr && x.Assessment_Result_Version == Assessment_Result_Version)
                    .AsEnumerable().Select(x => DbToD66ViewModel(x, _Assessor, D66Files)).ToList();
            }
            return datas;
        }
        #endregion

        #region get Assessment_Stage
        /// <summary>
        /// get Assessment_Stage
        /// </summary>
        /// <param name="Assessment_Kind"></param>
        /// <param name="Assessment_Sub_Kind"></param>
        /// <returns></returns>
        public List<string> getAssessmentStage(
            string Assessment_Kind,
            string Assessment_Sub_Kind
            )
        {
            List<string> result = new List<string>();
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                var data = db.Bond_Assessment_Parm.AsNoTracking()
                    .Where(x => x.Assessment_Kind == Assessment_Kind &&
                    x.IsActive == "Y" &&
                    x.Assessment_Sub_Kind == Assessment_Sub_Kind)
                    .Select(x => x.Assessment_Stage).Distinct();
                if (data.Any())
                    result = data.ToList();
            }
            return result;
        }
        #endregion

        #region get Assessment Sub Kind
        /// <summary>
        /// get Assessment Sub Kind
        /// </summary>
        /// <param name="type"></param>
        /// <returns></returns>
        public List<SelectOption> getAssessmentSubKind(AssessmentKind_Type type)
        {
            List<SelectOption> results = new List<SelectOption>() {
                new SelectOption() {Text = "All",Value = "All" } };
            var assessmentKindType = type.GetDescription();
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                results.AddRange(db.Bond_Assessment_Parm.AsNoTracking().Where(x =>
                x.IsActive == "Y" &&
                x.Assessment_Kind == assessmentKindType)
                .Select(x => x.Assessment_Sub_Kind).Distinct()
                .AsEnumerable().Select(x => getAssessmentSubKindOption(x)).ToList());
            }
            return results;
        }
        #endregion

        #region get Check_Item_Code
        /// <summary>
        /// get Check_Item_Code
        /// </summary>
        /// <param name="Assessment_Stage"></param>
        /// <param name="Assessment_Kind"></param>
        /// <param name="Assessment_Sub_Kind"></param>
        /// <returns></returns>
        public List<Tuple<string,string>> getCheckItemCode(
            string Assessment_Stage,
            string Assessment_Kind,
            string Assessment_Sub_Kind
            )
        {
            List<Tuple<string, string>> result = new List<Tuple<string, string>>();
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                try
                {
                    result = db.Bond_Assessment_Parm.AsNoTracking().Where(x =>
                             x.Assessment_Kind == Assessment_Kind &&
                             x.Assessment_Stage == Assessment_Stage &&
                             x.Assessment_Sub_Kind == Assessment_Sub_Kind &&
                             x.IsActive == "Y" &&
                             !x.Check_Item_Code.StartsWith(checkItemCode)).AsEnumerable()
                             .Select(x => new Tuple<string, string>(x.Check_Item_Code,
                             x.Check_Item_Memo == null ? string.Empty :
                             x.Check_Item_Memo.Replace("，", "，\n").Replace("；", "；\n"))).ToList();
                }
                catch (Exception ex)
                {

                }
            }
            return result;
        }
        #endregion

        #region get QuantifyFile fileName
        public List<Bond_Quantitative_Result_File> getQuantifyFile(string Check_Reference)
        {
            List<Bond_Quantitative_Result_File> result = new List<Bond_Quantitative_Result_File>();
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                result = db.Bond_Quantitative_Result_File.AsNoTracking()
                            .Where(x => x.Check_Reference == Check_Reference &&
                            x.Status == "Y").ToList();
            }
            return result;
        }
        #endregion

        #region get QualitativeFile fileName
        /// <summary>
        /// get QualitativeFile fileName
        /// </summary>
        /// <param name="Check_Reference"></param>
        /// <returns></returns>
        public List<Bond_Qualitative_Assessment_Result_File> getQualitativeFile(string Check_Reference)
        {
            List<Bond_Qualitative_Assessment_Result_File> result = new List<Bond_Qualitative_Assessment_Result_File>();
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                result =  db.Bond_Qualitative_Assessment_Result_File.AsNoTracking()
                            .Where(x => x.Check_Reference == Check_Reference &&
                            x.Status == "Y").ToList();
            }
            return result;
        }
        #endregion

        #region get Assessment_Result , Memo
        /// <summary>
        /// get Assessment_Result , Memo
        /// </summary>
        /// <param name="ReportDate"></param>
        /// <param name="Reference_Nbr"></param>
        /// <param name="Vresion"></param>
        /// <param name="Assessment_Result_Version"></param>
        /// <param name="type"></param>
        /// <param name="table"></param>
        /// <returns></returns>
        public string GetTextArea(
            DateTime ReportDate,
            string Reference_Nbr,
            int Vresion,
            int Assessment_Result_Version,
            string Check_Item_Code,
            string type,
            Table_Type table
            )
        {
            string result = string.Empty;
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                switch (table)
                {
                    case Table_Type.D63:
                        var D63 = db.Bond_Quantitative_Resource.AsNoTracking()
                            .FirstOrDefault(x => x.Reference_Nbr == Reference_Nbr &&
                            x.Report_Date == ReportDate &&
                            x.Version == Vresion &&
                            x.Assessment_Result_Version == Assessment_Result_Version);
                        if (D63 != null)
                        {
                            if (type == "Auditor_Reply")
                                result = D63.Auditor_Reply;
                        }
                        break;
                    case Table_Type.D64:
                        var D64 = db.Bond_Quantitative_Result.AsNoTracking()
                             .FirstOrDefault(x => x.Reference_Nbr == Reference_Nbr &&
                            x.Report_Date == ReportDate &&
                            x.Version == Vresion &&
                            x.Assessment_Result_Version == Assessment_Result_Version &&
                            x.Check_Item_Code == Check_Item_Code);
                        if (D64 != null)
                        {
                            if (type == "Assessment_Result")
                                result = D64.Assessment_Result;
                            if (type == "Memo")
                                result = D64.Memo;
                        }
                        break;
                    case Table_Type.D65:
                        var D65 = db.Bond_Qualitative_Assessment.AsNoTracking()
                            .FirstOrDefault(x => x.Reference_Nbr == Reference_Nbr &&
                            x.Report_Date == ReportDate &&
                            x.Version == Vresion &&
                            x.Assessment_Result_Version == Assessment_Result_Version);
                        if (D65 != null)
                        {
                            if (type == "Auditor_Reply")
                                result = D65.Auditor_Reply;
                        }
                        break;
                    case Table_Type.D66:
                        var D66 = db.Bond_Qualitative_Assessment_Result.AsNoTracking()
                             .FirstOrDefault(x => x.Reference_Nbr == Reference_Nbr &&
                            x.Report_Date == ReportDate &&
                            x.Version == Vresion &&
                            x.Assessment_Result_Version == Assessment_Result_Version &&
                            x.Check_Item_Code == Check_Item_Code);
                        if (D66 != null)
                        {
                            if (type == "Assessment_Result")
                                result = D66.Assessment_Result;
                            if (type == "Memo")
                                result = D66.Memo;
                        }
                        break;
                }             
            }
            return result;
        }
        #endregion

        #region get 減損試算結果
        /// <summary>
        /// get 減損試算結果
        /// </summary>
        /// <param name="Reference_Nbr"></param>
        /// <returns></returns>
        public List<ChgInSpreadViewModel> getChgInSpread(string Reference_Nbr)
        {
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                List<ChgInSpreadViewModel> results = new List<ChgInSpreadViewModel>();
                var A41 = db.Bond_Account_Info.AsNoTracking().FirstOrDefault(x => x.Reference_Nbr == Reference_Nbr);
                var C07s = db.EL_Data_Out.AsNoTracking().Where(x => x.Reference_Nbr == Reference_Nbr);
                if (A41 == null || !C07s.Any())
                    return new List<ChgInSpreadViewModel>();
                C07s.ToList().ForEach(C07 =>
                {
                    ChgInSpreadViewModel result = new ChgInSpreadViewModel();
                    result.FLOWID = C07.FLOWID;
                    result.PRJID = C07.PRJID;
                    var Ex_rate = (A41.Ex_rate != null ? A41.Ex_rate.Value : 0);
                    var Ori_Ex_rate = (A41.Ori_Ex_rate != null ? A41.Ori_Ex_rate.Value : 0);
                    var Data_A = A41.Currency_Code; //以Reference_Nbr 整合A41 取得之Currency_Code
                    var Data_B = C07.Y1_EL; //C07.Y1_EL
                    var Data_C = Ex_rate * Data_B; //Ex_rate  *  B
                    var Data_D = Ori_Ex_rate * Data_B; //Ori_Ex_rate  * B
                    var Data_E = C07.Lifetime_EL; // Lifetime_EL
                    var Data_F = Ex_rate * Data_E; // Ex_rate  *  E
                    var Data_G = Ori_Ex_rate * Data_E; // Ori_Ex_rate  * E
                    var Data_H = C07.PD > 0 ? C07.Y1_EL / C07.PD : 0; //IF PD > 0 H=Y1_EL/PD IF PD = 0 H = 0
                    var Data_I = Ex_rate * Data_H; //Ex_rate  *  H
                    var Data_J = Ori_Ex_rate * Data_H; //Ori_Ex_rate  * H
                    result.ChgInSpread.Add(new ChgInSpreadDetail()
                    {
                        Currency_Code = Data_A,
                        Base = "原幣",
                        Stage_1 = Data_B.ToString(),
                        Stage_2 = Data_E.ToString(),
                        Stage_3 = Data_H.ToString()
                    });
                    result.ChgInSpread.Add(new ChgInSpreadDetail()
                    {
                        Currency_Code = "TWD</br>台幣",
                        Base = "報導日匯率",
                        Stage_1 = Data_C.ToString(),
                        Stage_2 = Data_F.ToString(),
                        Stage_3 = Data_I.ToString()
                    });
                    result.ChgInSpread.Add(new ChgInSpreadDetail()
                    {
                        //Currency_Code = "台幣",
                        Base = "成本匯率",
                        Stage_1 = Data_D.ToString(),
                        Stage_2 = Data_G.ToString(),
                        Stage_3 = Data_J.ToString()
                    });
                    results.Add(result);
                });
                return results;
            }
        }
        #endregion

        #region  get 減損作業狀態查詢 
        /// <summary>
        /// get 減損作業狀態查詢 
        /// </summary>
        /// <param name="reportDate"></param>
        /// <returns></returns>
        public List<D6CheckViewModel> getD6Check(string reportDate)
        {
            List<D6CheckViewModel> result = new List<D6CheckViewModel>();
            DateTime dt = DateTime.MinValue;
            if (!DateTime.TryParse(reportDate, out dt))
            {
                return result;
            }
            var _completed = Stage_Type.completed.GetDescription();
            var _undone = Stage_Type.undone.GetDescription();
            var _NoneAssess = Stage_Type.No_Need_Assess.GetDescription();
            var _NoneReview = Stage_Type.No_Need_Review.GetDescription();
            var C07ReportDate = dt.ToString("yyyy-MM-dd");
            var _dtStr = dt.ToString("yyyy/MM/dd");
            var _WatchFlag = false; //有無觀察名單
            var _D65Flag = false; //有無進行質化評估(true:需進行質化評估)
            var _UuDoneFlag = false; //後續是否要繼續
            List<Bond_Quantitative_Resource> D63s = new List<Bond_Quantitative_Resource>(); //D63
            List<string> _D63Refs = new List<string>(); //D63 Refs
            List<Bond_Qualitative_Assessment> D65s = new List<Bond_Qualitative_Assessment>(); //D65
            List<string> _D65Refs = new List<string>(); //D65 Refs
            var _Detail = string.Empty;
            var _Job_Details = string.Empty;
            var _lastVersion = 0;
            StringBuilder sb = new StringBuilder();
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                #region 減損計算
                sb = new StringBuilder();
                _Job_Details = Impairment_Operations_Type.ImpairmentCalculation.GetDescription();
                var C07s = db.EL_Data_Out.AsNoTracking().Where(x => x.Report_Date == C07ReportDate).ToList();
                if (C07s.Any())
                {
                    sb.AppendLine($@"報導日:{_dtStr}");
                    foreach (var item in C07s.GroupBy(x => new { x.PRJID, x.FLOWID }))
                    {
                        var version = item.Max(y => y.Version);
                        var _count = item.Where(y => y.Version == version).Count();
                        sb.AppendLine($@"專案:{item.Key.PRJID} 流程:{item.Key.FLOWID} 最大版本:{version} 資料筆數:{_count}");
                    }
                    AddD6CheckView(result, _Job_Details, _completed, sb); //減損計算:已完成
                }
                else
                {
                    _UuDoneFlag = true;
                    AddD6CheckView(result, _Job_Details, _undone, sb); //減損計算:未完成
                }
                #endregion
                #region 基本要件及監控名單產出
                sb = new StringBuilder();
                _Job_Details = Impairment_Operations_Type.BasicRequirements.GetDescription();
                if (_UuDoneFlag)
                {
                    AddD6CheckView(result, _Job_Details, _undone, sb); //基本要件及監控名單產出:未完成
                }
                else
                {
                    var D62s = db.Bond_Basic_Assessment.AsNoTracking().Where(x => x.Report_Date == dt).ToList();
                    if (D62s.Any())
                    {
                        _lastVersion = D62s.Max(x => x.Version);
                        var _D62s = D62s.Where(x => x.Version == _lastVersion).ToList();
                        var _watching = _D62s.Where(x => x.Watch_IND == "Y").Count();
                        sb.AppendLine($@"報導日:{_dtStr} 資料版本:{_lastVersion} 資料處理日期:{_D62s.First().Processing_Date.ToString("yyyy/MM/dd")}");
                        sb.AppendLine($@"總資料數:{_D62s.Count()} 觀察名單筆數:{_watching} 預警名單筆數:{_D62s.Where(x => x.Warning_IND == "Y").Count()}");
                        if (_watching > 0)
                            _WatchFlag = true;
                        AddD6CheckView(result, _Job_Details, _completed, sb); //基本要件及監控名單產出:已完成
                    }
                    else
                    {
                        _UuDoneFlag = true;
                        AddD6CheckView(result, _Job_Details, _undone, sb); //基本要件及監控名單產出:未完成
                    }
                }
                #endregion
                #region 是否完成量化評估
                sb = new StringBuilder();
                _Job_Details = Impairment_Operations_Type.QuantitativeAssessment.GetDescription();
                if (_UuDoneFlag)
                {
                    AddD6CheckView(result, _Job_Details, _undone, sb); //是否完成量化評估:未完成                   
                }
                else
                {
                    if (_WatchFlag) //有觀察名單
                    {
                        var confirmflag = true; //是否評估完成
                        D63s = db.Bond_Quantitative_Resource.AsNoTracking()
                            .Where(x => x.Report_Date == dt && x.Version == _lastVersion).ToList();
                        if (D63s.Any())
                        {
                            _D63Refs = D63s.GroupBy(x => x.Reference_Nbr).Select(x => x.Key).ToList(); //帳戶數
                            foreach (var Ref in _D63Refs)
                            {
                                //每一個帳戶都有一筆提交複核(Send_to_Auditor == "Y"),且無被退回(Auditor_Return == null)
                                if (!D63s.Any(x => x.Reference_Nbr == Ref && x.Send_to_Auditor == "Y" && x.Auditor_Return == null))
                                {
                                    confirmflag = false;
                                    break;
                                }
                            }
                        }
                        else
                        {
                            confirmflag = false;
                        }
                        if (confirmflag)
                        {
                            sb.AppendLine($@"報導日:{_dtStr} 最大評估日期:{D63s.Max(x => x.Assessment_date)?.ToString("yyyy/MM/dd")} 帳戶數:{_D63Refs.Count}");
                            AddD6CheckView(result, _Job_Details, _completed, sb); //是否完成量化評估:已完成
                        }
                        else
                        {
                            _UuDoneFlag = true;
                            AddD6CheckView(result, _Job_Details, _undone, sb); //是否完成量化評估:未完成
                        }
                    }
                    else
                    {
                        AddD6CheckView(result, _Job_Details, _NoneAssess, sb); //是否完成量化評估:無須評估
                    }
                }
                #endregion
                #region 是否完成量化評估複核
                sb = new StringBuilder();
                _Job_Details = Impairment_Operations_Type.QuantitativeAssessmentReview.GetDescription();
                if (_UuDoneFlag)
                {
                    AddD6CheckView(result, _Job_Details, _undone, sb); //是否完成量化複核:未完成
                }
                else
                {
                    if (_WatchFlag) //有觀察名單
                    {
                        var confirmflag = true; //是否複核完成
                        var _AuditorCount = 0; //完成複核數
                        if (D63s.Any())
                        {
                            foreach (var Ref in _D63Refs)
                            {
                                //每一個帳戶都有一筆複核完成(Assessment_Result_Version == Result_Version_Confirm)
                                var _D63 = D63s.FirstOrDefault(x => x.Reference_Nbr == Ref && x.Assessment_Result_Version == x.Result_Version_Confirm);
                                if (_D63 != null)
                                {
                                    _AuditorCount += 1;
                                    //複核後狀態不通過(Quantitative_Pass_Confirm == "N")
                                    if (_D63.Quantitative_Pass_Confirm == "N")
                                        _D65Flag = true;
                                }
                                else
                                {
                                    confirmflag = false;
                                    break;
                                }
                            }
                            sb.AppendLine($@"報導日:{_dtStr} 最大複核日期:{D63s.Max(x => x.Audit_date)?.ToString("yyyy/MM/dd")} 帳戶數:{_D63Refs.Count} 完成複核帳戶數:{_AuditorCount}");
                        }
                        else
                        {
                            confirmflag = false;
                        }
                        if (confirmflag)
                        {
                            AddD6CheckView(result, _Job_Details, _completed, sb); //是否完成量化複核:已完成
                        }
                        else
                        {
                            _UuDoneFlag = true;
                            AddD6CheckView(result, _Job_Details, _undone, sb); //是否完成量化複核:未完成
                        }
                    }
                    else
                    {
                        AddD6CheckView(result, _Job_Details, _NoneReview, sb); //是否完成量化複核:無須複核
                    }
                }
                #endregion
                #region 是否完成質化評估
                sb = new StringBuilder();
                _Job_Details = Impairment_Operations_Type.QualitativeAssessment.GetDescription();
                if (_UuDoneFlag)
                {
                    AddD6CheckView(result, _Job_Details, _undone, sb); //是否完成質化評估:未完成
                }
                else
                {
                    if (_WatchFlag && _D65Flag) //有觀察名單 & 需進行質化評估
                    {
                        var confirmflag = true; //是否評估完成
                        D65s = db.Bond_Qualitative_Assessment.AsNoTracking()
                            .Where(x => x.Report_Date == dt && x.Version == _lastVersion && x.Extra_Case != "D").ToList();
                        if (D65s.Any())
                        {
                            _D65Refs = D65s.GroupBy(x => x.Reference_Nbr).Select(x => x.Key).ToList(); //帳戶數
                            foreach (var Ref in _D65Refs)
                            {
                                //每一個帳戶都有一筆提交複核(Send_to_Auditor == "Y"),且無被退回(Auditor_Return == null)
                                if (!D65s.Any(x => x.Reference_Nbr == Ref && x.Send_to_Auditor == "Y" && x.Auditor_Return == null))
                                {
                                    confirmflag = false;
                                    break;
                                }
                            }
                        }
                        else
                        {
                            confirmflag = false;
                        }
                        if (confirmflag)
                        {
                            sb.AppendLine($@"報導日:{_dtStr} 最大評估日期:{D65s.Max(x => x.Assessment_date)?.ToString("yyyy/MM/dd")} 帳戶數:{_D65Refs.Count}");
                            AddD6CheckView(result, _Job_Details, _completed, sb); //是否完成質化評估:已完成
                        }
                        else
                        {
                            _UuDoneFlag = true;
                            AddD6CheckView(result, _Job_Details, _undone, sb); //是否完成質化評估:未完成
                        }
                    }
                    else
                    {
                        AddD6CheckView(result, _Job_Details, _NoneAssess, sb); //是否完成質化評估:無須評估
                    }
                }
                #endregion
                #region 是否完成質化評估複核
                sb = new StringBuilder();
                _Job_Details = Impairment_Operations_Type.QualitativeAssessmentReview.GetDescription();
                if (_UuDoneFlag)
                {
                    AddD6CheckView(result, _Job_Details, _undone, sb); //是否完成質化複核:未完成
                }
                else
                {
                    if (_WatchFlag && _D65Flag) //有觀察名單 & 需進行質化評估
                    {
                        var confirmflag = true; //是否複核完成
                        var _AuditorCount = 0; //完成複核數
                        if (D65s.Any())
                        {
                            foreach (var Ref in _D65Refs)
                            {
                                //每一個帳戶都有一筆複核完成(Assessment_Result_Version == Result_Version_Confirm)
                                var _D65 = D65s.FirstOrDefault(x => x.Reference_Nbr == Ref && x.Assessment_Result_Version == x.Result_Version_Confirm);
                                if (_D65 != null)
                                {
                                    _AuditorCount += 1;
                                }
                                else
                                {
                                    confirmflag = false;
                                    break;
                                }
                            }
                            sb.AppendLine($@"報導日:{_dtStr} 最大複核日期:{D65s.Max(x => x.Audit_date)?.ToString("yyyy/MM/dd")} 帳戶數:{_D65Refs.Count} 完成複核帳戶數:{_AuditorCount}");
                        }
                        else
                        {
                            confirmflag = false;
                        }
                        if (confirmflag)
                        {
                            AddD6CheckView(result, _Job_Details, _completed, sb); //是否完成質化複核:已完成
                        }
                        else
                        {
                            _UuDoneFlag = true;
                            AddD6CheckView(result, _Job_Details, _undone, sb); //是否完成質化複核:未完成
                        }
                    }
                    else
                    {
                        AddD6CheckView(result, _Job_Details, _NoneReview, sb); //是否完成質化複核:無須評估
                    }
                }
                #endregion
                #region 階段調整確認作業
                sb = new StringBuilder();
                _Job_Details = Impairment_Operations_Type.StageConfrim.GetDescription();
                if (_UuDoneFlag)
                {
                    AddD6CheckView(result, _Job_Details, _undone, sb); //階段調整確認作業:未完成
                }
                else
                {
                    var D54s = db.IFRS9_EL.AsNoTracking().Where(x => x.Report_Date == C07ReportDate && x.Version == _lastVersion).ToList();
                    if (D54s.Any())
                    {
                        sb.AppendLine($@"報導日:{_dtStr}");
                        foreach (var item in D54s.GroupBy(x => new { x.PRJID, x.FLOWID }))
                        {
                            var Exec_Date = item.First().Exec_Date?.ToString("yyyy/MM/dd");
                            var Impairment_Stage1 = item.Count(x => x.Impairment_Stage == "1");
                            var Impairment_Stage2 = item.Count(x => x.Impairment_Stage == "2");
                            var Impairment_Stage3 = item.Count(x => x.Impairment_Stage == "3");
                            sb.AppendLine($@"專案:{item.Key.PRJID} 流程:{item.Key.FLOWID} ");
                            sb.AppendLine($@"資料版本:{_lastVersion} 資料處理日期:{Exec_Date}");
                            sb.AppendLine($@"Stage1資料筆數:{Impairment_Stage1} Stage2資料筆數:{Impairment_Stage2} Stage3資料筆數:{Impairment_Stage3}");
                        }
                        AddD6CheckView(result, _Job_Details, _completed, sb); //階段調整確認作業:已完成
                    }
                    else
                    {
                        _UuDoneFlag = true;
                        AddD6CheckView(result, _Job_Details, _undone, sb); //階段調整確認作業:未完成
                    }
                }
                #endregion
            }
            return result;
        }
        #endregion

        #region 抓取A41 該基準日最大版本
        /// <summary>
        /// 抓取A41 該基準日最大版本
        /// </summary>
        /// <param name="reportDate"></param>
        /// <returns></returns>
        public int GetA41Version(string reportDate)
        {
            int ver = 0;
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                DateTime _reportDate = DateTime.Parse(reportDate);
                var data = db.Bond_Account_Info.AsNoTracking()
                    .Where(x => x.Report_Date == _reportDate)
                    .Select(x => x.Version).Distinct().ToList();
                if (data.Any())
                    ver = data.Max(x => x).Value;
            }
            return ver;
        }
        #endregion
        #region 抓取A41 該基準日Assessment_Check為N的部位(190619 PGE需求新增)
        public List<Bond_Account_Info> GetA41AssessmentCheck(string reportDate, int version)
        {
            List<Bond_Account_Info> A41AssessmentCheck = new List<Bond_Account_Info>();
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                DateTime _reportDate = DateTime.Parse(reportDate);
                A41AssessmentCheck = db.Bond_Account_Info.AsNoTracking()
                    .Where(x => x.Report_Date == _reportDate && x.Version == version && (x.Assessment_Check == "N" || x.Assessment_Check == "n"))
                    .ToList();

            }
            return A41AssessmentCheck;
        }
        #endregion
        #region get EL FLOW_ID

        /// <summary>
        /// get EL FLOW_ID
        /// </summary>
        /// <param name="reportDate"></param>
        /// <returns></returns>
        public List<string> getFlowIDs(string reportDate)
        {
            List<string> result = new List<string>();
            DateTime dt = DateTime.MinValue;
            DateTime.TryParse(reportDate, out dt);
            var _reportDate = dt.ToString("yyyy-MM-dd");
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                result = db.IFRS9_EL.AsNoTracking()
                    .Where(x => x.Report_Date == _reportDate)
                    .Select(x => x.FLOWID)
                    .Distinct()
                    .AsEnumerable()
                    .ToList();
            }
            return result;
        }

        /// <summary>
        /// 查詢 ReEL
        /// </summary>
        /// <param name="reportDate"></param>
        /// <returns></returns>
        public List<ReELViewModel> getReEL(string reportDate ,string Group_Product_Code)
        {
            List<ReELViewModel> result = new List<ReELViewModel>();
            string _reportDate = string.Empty;
            DateTime dt = DateTime.MinValue;
            DateTime.TryParse(reportDate, out dt);
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                var users = db.IFRS9_User.AsNoTracking().ToList();
                var flowIds = db.Flow_Info.AsNoTracking()
                    .Where(x => x.Group_Product_Code.StartsWith("4"))
                    .Where(x => x.Group_Product_Code == Group_Product_Code, !Group_Product_Code.IsNullOrWhiteSpace())
                    .Select(x => x.FLOWID).ToList();
                result = db.IFRS9_RE_EL.AsNoTracking()
                    .Where(x => x.Report_Date == dt, dt != DateTime.MinValue)
                    .Where(x => flowIds.Contains(x.FLOWID))
                    .OrderByDescending(x => x.Report_Date)
                    .ThenBy(x => x.FLOWID)
                    .ThenByDescending(x => x.Create_Date)
                    .ThenByDescending(x => x.Create_Time)
                    .AsEnumerable()
                    .Select(x => new ReELViewModel()
                    {
                        Report_Date = x.Report_Date.ToString("yyyy/MM/dd"),
                        FLOWID = x.FLOWID,
                        Message = x.Message,
                        Create_User = $"{x.Create_User}({users.FirstOrDefault(y => y.User_Account == x.Create_User)?.User_Name})",
                        Create_Time = $"{x.Create_Date.ToString("yyyy/MM/dd")} {x.Create_Time.timeSpanToStr()}"
                    }).ToList();
            }
            return result;
        }

        #endregion

        #endregion Get Data

        #region Save DB 部分

        #region 執行量化評估需求
        /// <summary>
        /// A41 data to D63
        /// </summary>
        /// <param name="reportDate"></param>
        /// <param name="version"></param>
        /// <param name="bondNumber"></param>
        /// <returns></returns>
        public MSGReturnModel TransferD63(string reportDate)
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            DateTime startDt = DateTime.Now;
            DateTime dt = DateTime.MinValue;
            int version = 0;
            if (!DateTime.TryParse(reportDate, out dt))
            {
                result.DESCRIPTION = Message_Type.parameter_Error.GetDescription();
                return result;
            }
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {               
                if (getD54Check(dt))
                {
                    result.DESCRIPTION = Message_Type.already_Save_D54.GetDescription($"{dt.ToString("yyyyMMdd")}資料");
                    return result;
                }
              
                var D62 = db.Bond_Basic_Assessment.AsNoTracking().Where(x => x.Report_Date == dt).ToList();
                if (D62.Any())
                {
                    version = D62.Max(x => x.Version);

                    var _cr = $"{ dt.ToString("yyyyMMdd")}_V{version}";

                    var D64Files = db.Bond_Quantitative_Result_File
                        .Where(x => x.Check_Reference.StartsWith(_cr)).ToList();

                    var D66Files = db.Bond_Qualitative_Assessment_Result_File
                        .Where(x => x.Check_Reference.StartsWith(_cr)).ToList();

                    D64Files.ForEach(x =>
                    {
                        try
                        {
                            File.Delete(x.File_path);
                        }
                        catch
                        {

                        }
                    });

                    D66Files.ForEach(x =>
                    {
                        try
                        {
                            File.Delete(x.File_path);
                        }
                        catch
                        {

                        }
                    });

                    db.Bond_Quantitative_Result_File.RemoveRange(D64Files);
                    db.Bond_Qualitative_Assessment_Result_File.RemoveRange(D66Files);
                
                    if (!D62.Any(x => x.Watch_IND == "Y" && x.Version == version))
                    {
                        string delSql = string.Empty;

                        delSql = $@"
DELETE Bond_Quantitative_Resource
WHERE Report_Date = @Report_Date
AND Version = @Version ;

DELETE Bond_Quantitative_Result
WHERE Report_Date = @Report_Date
AND Version = @Version ;

DELETE Bond_Qualitative_Assessment
WHERE Report_Date = @Report_Date
AND Version = @Version ;

DELETE Bond_Qualitative_Assessment_Result
WHERE Report_Date = @Report_Date
AND Version = @Version ; ";

                        try
                        {
                            int i = db.Database.ExecuteSqlCommand(delSql,
                                new List<SqlParameter>()
                                {
                                  new SqlParameter("Report_Date", dt.ToString("yyyy-MM-dd")),
                                  new SqlParameter("Version", version.ToString())
                                }.ToArray());
                        
                            db.SaveChanges();
                        }
                        catch
                        {

                        }
                        result.DESCRIPTION = "D62最大版本沒有符合觀察名單(需進行量化)的資料";
                        return result;
                    }
                }
                else
                {
                    result.DESCRIPTION = Message_Type.not_Find_Any.GetDescription("D62", "無法執行!");
                    return result;
                }

                #region Sql
                string sql = string.Empty;
                var dtStrSql = dt.dateTimeToStrSql();
                var versionSql = version.intToStrSql();
                var startDtSql = startDt.dateTimeToStrSql();
                var otherStr = AssessmentSubKind_Type.Other.GetDescription();
                var SGDStr = AssessmentSubKind_Type.Sovereign_Government_Debt.GetDescription();
                var AbuDhabiStr = "UNITED ARAB EMI";
                sql = $@"
DELETE Bond_Quantitative_Resource
WHERE Report_Date = {dtStrSql}
AND Version = {versionSql} ;

DELETE Bond_Quantitative_Result
WHERE Report_Date = {dtStrSql}
AND Version = {versionSql} ;

DELETE Bond_Qualitative_Assessment
WHERE Report_Date = {dtStrSql}
AND Version = {versionSql} ;

DELETE Bond_Qualitative_Assessment_Result
WHERE Report_Date = {dtStrSql}
AND Version = {versionSql} ;


--D63 
WITH
TEMPA46s AS
(
   SELECT *
   FROM Fixed_Income_CEIC_Info
   Where Effective = 'Y'
),
TEMPA47s AS
(
   SELECT *
   FROM Fixed_Income_Moody_Info
   Where Effective = 'Y'
),
TEMPA48s AS
(
   SELECT *
   FROM Fixed_Income_AbuDhabi_Info
   Where Effective = 'Y'
   AND   Data_Year >= DATEPART(YEAR,{dtStrSql})-4
),
TEMPA91s AS
(
   SELECT *
   FROM Gov_Info_Yearly
   WHERE Data_Year >= DATEPART(YEAR,{dtStrSql})-4
),
TEMPA92s AS
(
   SELECT *
   FROM Gov_Info_Quartly
),
TEMPA93s AS
(
   SELECT *
   FROM Gov_Info_Monthly
),
TEMPA48Year AS
(
	 SELECT *
	 FROM TEMPA48s  
	 where Data_Year < DATEPART(YEAR, {dtStrSql} )
),
TEMPA48 AS
(
   SELECT top 1 *
   FROM TEMPA48Year
   order by Data_Year DESC
),
TEMPA91Year AS
(
	 SELECT *
	 FROM TEMPA91s  
	 where Data_Year < DATEPART(YEAR, {dtStrSql} )
),
TEMPA91 AS 
(
   SELECT *
   FROM TEMPA91s
   WHERE Data_Year = DATEPART(YEAR,{dtStrSql})
)
INSERT INTO [Bond_Quantitative_Resource]
           ([Reference_Nbr]
           ,[Bond_Number]
           ,[Lots]
           ,[Segment_Name]
           ,[Version]
           ,[Report_Date]
           ,[Processing_Date]
           ,[Assessment_Sub_Kind]
           ,[Net_Debt] 
           ,[Total_Asset] 
           ,[CFO] 
           ,[Int_Expense] 
           ,[Total_Equity] 
           ,[BS_TOT_ASSET] 
           ,[SHORT_AND_LONG_TERM_DEBT] 
           ,[ISSUER_AREA]
           ,[IGS_INDEX]
           ,[IGS_INDEX_Increase]
           ,[External_Debt_Rate]
           ,[FC_Indexed_Debt_Rate]
           ,[CEIC_Value]
           ,[Short_term_Debt_Ratio]
           ,[GDP_Yearly_Avg]
           ,[Portfolio]
           ,[Origination_Date]
           ,[Portfolio_Name]
           ,[Assessment_Result_Version]
           ,[Result_Version_Confirm]
           ,[Quantitative_Pass_Confirm]
           ,[Create_User]
           ,[Create_Date]
           ,[Create_Time]
)
SELECT 
A41.Reference_Nbr AS Reference_Nbr, --帳戶編號/群組編號
A41.Bond_Number AS Bond_Number, --債券編號
A41.Lots AS Lots, --Lots
A41.Segment_Name AS Segment_Name, --債券(資產)名稱
A41.Version AS Version, --資料版本
A41.Report_Date AS Report_Date, --評估基準日/報導日
{startDtSql} AS Processing_Date, -- 資料處理日期
A41.Assessment_Sub_Kind AS Assessment_Sub_Kind, --評估次分類
A53_SampleInfo.Net_Debt AS Net_Debt, 
A53_SampleInfo.Total_Asset / CAST(1000000 AS float)  AS Total_Asset,
A53_SampleInfo.CFO AS CFO,
A53_SampleInfo.Int_Expense AS Int_Expense,
A53_SampleInfo.Total_Equity / CAST(1000000 AS float)  AS Total_Equity,
A53_SampleInfo.BS_TOT_ASSET / CAST(1000000 AS float)  AS BS_TOT_ASSET,
A53_SampleInfo.SHORT_AND_LONG_TERM_DEBT AS SHORT_AND_LONG_TERM_DEBT,
A41.ISSUER_AREA AS ISSUER_AREA, --Issuer所屬區域
CASE WHEN A41.Assessment_Sub_Kind ='{SGDStr}' AND A41.ISSUER_AREA <> '{AbuDhabiStr}'
     THEN (
	 SELECT top 1 IGS_INDEX from TEMPA91Year where Country = A41.ISSUER_AREA
     order by  Data_Year DESC )
     WHEN A41.Assessment_Sub_Kind ='{SGDStr}' AND A41.ISSUER_AREA = '{AbuDhabiStr}'
	 THEN (
	 SELECT IGS_INDEX from TEMPA48)
END  AS IGS_INDEX, --政府債務/GDP
CASE WHEN A41.Assessment_Sub_Kind ='{SGDStr}' AND A41.ISSUER_AREA <> '{AbuDhabiStr}'
     THEN (
	 (select top 1 IGS_INDEX from TEMPA91Year where Country = A41.ISSUER_AREA
     order by  Data_Year DESC )
	 -
	 (select Min(IGS_INDEX) from TEMPA91Year where Country = A41.ISSUER_AREA)
     )
	 WHEN A41.Assessment_Sub_Kind ='{SGDStr}' AND A41.ISSUER_AREA = '{AbuDhabiStr}'
	 THEN (
     (SELECT IGS_INDEX from TEMPA48)
     -
     (Select IGS_INDEX from TEMPA48Year  
     where Data_Year = DATEPART(YEAR,{dtStrSql})-4 )
     )
END AS IGS_INDEX_Increase, --過去4年政府債務/GDP增加數
CASE WHEN A41.Assessment_Sub_Kind ='{SGDStr}' AND A41.ISSUER_AREA <> '{AbuDhabiStr}'
     THEN (
	 SELECT TOP 1 External_Debt_Rate
	 FROM TEMPA47s A47
	 WHERE A47.Country = A41.ISSUER_AREA
     order by A47.Processing_Date desc)
	 WHEN A41.Assessment_Sub_Kind ='{SGDStr}' AND A41.ISSUER_AREA = '{AbuDhabiStr}'
	 THEN (
	 SELECT TOP 1 External_Debt_Rate
	 FROM TEMPA48 A48
	 WHERE A48.Country = A41.ISSUER_AREA)
END AS External_Debt_Rate, --外人持有政府債務/總政府債務
CASE WHEN A41.Assessment_Sub_Kind ='{SGDStr}' AND A41.ISSUER_AREA <> '{AbuDhabiStr}'
     THEN (
	 SELECT TOP 1 FC_Indexed_Debt_Rate
	 FROM TEMPA47s A47
	 WHERE A47.Country = A41.ISSUER_AREA
     order by A47.Processing_Date desc)
	 WHEN A41.Assessment_Sub_Kind ='{SGDStr}' AND A41.ISSUER_AREA = '{AbuDhabiStr}'
	 THEN (
	 SELECT TOP 1 FC_Indexed_Debt_Rate
	 FROM TEMPA48 A48
	 WHERE A48.Country = A41.ISSUER_AREA)
END AS FC_Indexed_Debt_Rate, --外幣計價政府債務/總政府債務
CASE WHEN A41.Assessment_Sub_Kind ='{SGDStr}' AND A41.ISSUER_AREA <> '{AbuDhabiStr}'
     THEN (
	 SELECT TOP 1 CEIC_Value
	 FROM TEMPA46s A46
	 WHERE A46.Country = A41.ISSUER_AREA
     order by A46.Processing_Date desc)
	 WHEN A41.Assessment_Sub_Kind ='{SGDStr}' AND A41.ISSUER_AREA = '{AbuDhabiStr}'
	 THEN (
	 SELECT TOP 1 CEIC_Value
	 FROM TEMPA48 A48
	 WHERE A48.Country = A41.ISSUER_AREA)
END AS CEIC_Value, --(經常帳+FDI淨流入)/GDP
CASE WHEN A41.Assessment_Sub_Kind ='{SGDStr}' AND A41.ISSUER_AREA <> '{AbuDhabiStr}'
     THEN (
	 (SELECT TOP 1 Short_term_Debt
	 FROM TEMPA92s A92
	 WHERE A92.Country = A41.ISSUER_AREA
	 ORDER BY A92.Processing_Date DESC) /
	 (CASE WHEN (SELECT TOP 1 Foreign_Exchange
	 FROM TEMPA93s A93
	 WHERE A93.Country = A41.ISSUER_AREA
	 ORDER BY A93.Processing_Date DESC) = 0
	 THEN null
	 ELSE (SELECT TOP 1 Foreign_Exchange
	 FROM TEMPA93s A93
	 WHERE A93.Country = A41.ISSUER_AREA
	 ORDER BY A93.Processing_Date DESC) END))
	 WHEN A41.Assessment_Sub_Kind ='{SGDStr}' AND A41.ISSUER_AREA = '{AbuDhabiStr}'
	 THEN (
	 SELECT (A48.Short_term_Debt / 
	 CASE WHEN A48.Foreign_Exchange = 0 THEN null ELSE A48.Foreign_Exchange END)
	 FROM TEMPA48 A48 
	 WHERE A48.Country = A41.ISSUER_AREA)
END AS Short_term_Debt_Ratio, --短期外債/外匯儲備
CASE WHEN A41.Assessment_Sub_Kind ='{SGDStr}' AND A41.ISSUER_AREA <> '{AbuDhabiStr}'
     THEN (
     SELECT Avg(A91.GDP_Yearly) 
     FROM  TEMPA91Year A91
     WHERE A91.Country = A41.ISSUER_AREA
     GROUP BY A91.Country)
	 WHEN A41.Assessment_Sub_Kind ='{SGDStr}' AND A41.ISSUER_AREA = '{AbuDhabiStr}'
	 THEN (
     SELECT Avg(A48.GDP_Yearly) 
     FROM  TEMPA48Year A48
     WHERE A48.Country = A41.ISSUER_AREA
     GROUP BY A48.Country)
END AS GDP_Yearly_Avg, -- 過去四年平均經濟成長
A41.Portfolio AS Portfolio, --Portfolio
A41.Origination_Date AS Origination_Date, --債券購入(認列)日期
A41.Portfolio_Name AS Portfolio_Name, --Portfolio英文
1, --Assessment_Result_Version
CASE WHEN A41.Assessment_Sub_Kind = '{otherStr}'
     THEN 2
     ELSE null
END AS Result_Version_Confirm,
CASE WHEN A41.Assessment_Sub_Kind = '{otherStr}'
     THEN 'N'
     ELSE null
END AS Quantitative_Pass_Confirm,
{_UserInfo._user.stringToStrSql()} AS Create_User,
{_UserInfo._date.dateTimeToStrSql()} AS Create_Date,
{_UserInfo._time.timeSpanToStrSql()} AS Create_Time
FROM 
(
Select * from Bond_Basic_Assessment
WHERE Report_Date = {dtStrSql}
AND Version = {versionSql}
AND Watch_IND = 'Y' 
) AS D62
JOIN 
(
Select * from Bond_Account_Info
WHERE Report_Date = {dtStrSql}
AND Version = {versionSql}
) AS A41
ON D62.Reference_Nbr = A41.Reference_Nbr
LEFT JOIN 
(
Select * from Rating_Info_SampleInfo
where Report_Date = {dtStrSql}
)
AS  A53_SampleInfo
ON  A41.Bond_Number = A53_SampleInfo.Bond_Number ;

--Add D64 Other
With Temp AS
(
select 
D63.Reference_Nbr,
D63.Report_Date,
D63.Version,
D61.Assessment_Kind,
D61.Assessment_Sub_Kind,
D61.Assessment_Stage,
D61.Check_Item_Code,
D61.Check_Item,
D61.Check_Item_Memo,
convert(varchar, D63.Report_Date, 112) + '_V' +
convert(varchar, D63.Version) + '_D63.Reference_Nbr_'+
'_ARV2_' +  D61.Check_Item_Code AS Check_Reference,
2 AS Assessment_Result_Version,
0 AS Check_Item_Value,
0 AS Pass_Value,
'N' AS Quantitative_Pass,
D61.Check_Symbol,
D61.Threshold,
null AS Memo,
D63.Bond_Number,
D63.Lots,
D63.Portfolio,
D63.Origination_Date,
D63.Portfolio_Name,
{"System".stringToStrSql()} AS Create_User,
{_UserInfo._date.dateTimeToStrSql()} AS Create_Date,
{_UserInfo._time.timeSpanToStrSql()} AS Create_Time
from 
(
   select * from Bond_Quantitative_Resource
   where Report_Date = {dtStrSql}
   and Version = {versionSql}
   and Assessment_Sub_Kind = '{otherStr}'
)
AS D63 --D63
join Bond_Assessment_Parm D61
ON D61.Check_Item_Code like '%X7%'
and D61.Assessment_Sub_Kind = D63.Assessment_Sub_Kind
and D61.Assessment_Kind = '量化衡量'
and D61.IsActive = 'Y'
)
INSERT INTO [Bond_Quantitative_Result]
           ([Reference_Nbr]
           ,[Report_Date] 
           ,[Version]
           ,[Assessment_Kind]
           ,[Assessment_Sub_Kind]
           ,[Assessment_Stage]
           ,[Check_Item_Code]
           ,[Check_Item]
           ,[Check_Item_Memo]
           ,[Check_Reference]
           ,[Assessment_Result_Version]
           ,[Check_Item_Value]
           ,[Pass_Value]
           ,[Quantitative_Pass]
           ,[Check_Symbol]
           ,[Threshold]
           ,[Memo]
           ,[Bond_Number]
           ,[Lots]
           ,[Portfolio]
           ,[Origination_Date]
           ,[Portfolio_Name]
           ,[Create_User]
           ,[Create_Date]
           ,[Create_Time]
)
select * from Temp;

--Add D63 Other Assessment_Result_Version = 2 
WITH TEMPD63Other AS(
SELECT [Reference_Nbr]
      ,[Bond_Number]
      ,[Lots]
      ,[Segment_Name]
      ,[Version]
      ,[Report_Date]
      ,[Processing_Date]
      ,[Assessment_Sub_Kind]
      ,[Net_Debt]
      ,[Total_Asset]
      ,[CFO]
      ,[Int_Expense]
      ,[CET1_Capital_Ratio]
      ,[Lerverage_Ratio]
      ,[Liquiduty_Coverage_Ratio]
      ,[Total_Equity]
      ,[BS_TOT_ASSET]
      ,[RBC_Ratio]
      ,[MCCSR_Ratio]
      ,[Solvency_II_Ratio]
      ,[SHORT_AND_LONG_TERM_DEBT]
      ,[Cash_and_Cash_Equivalents]
      ,[Short_Term_Debt]
      ,[ISSUER_AREA]
      ,[IGS_INDEX]
      ,[IGS_INDEX_Increase]
      ,[External_Debt_Rate]
      ,[FC_Indexed_Debt_Rate]
      ,[CEIC_Value]
      ,[Short_term_Debt_Ratio]
      ,[GDP_Yearly_Avg]
      ,[Quantitative_CLO_1]
      ,[Quantitative_CLO_2]
      ,[Quantitative_CLO_2_Threshold]
      ,[Quantitative_CLO_3]
      ,[Quantitative_CLO_3_Threshold]
      ,[Quantitative_CLO_4]
      ,[Quantitative_CLO_4_Threshold]
      ,[Portfolio]
      ,[Origination_Date]
      ,[Portfolio_Name]
      ,2 AS [Assessment_Result_Version]
      ,'System' AS [Assessor]
      ,{startDtSql} AS [Assessment_date]
      ,'System' AS [Auditor]
      ,{startDtSql} AS [Audit_date]
      ,'Y' AS [Send_to_Auditor]
      ,{_UserInfo._date.dateTimeToStrSql()} AS [Send_Time]
      ,2 AS [Result_Version_Confirm]
      ,'N' AS [Quantitative_Pass_Confirm]
      ,{"System".stringToStrSql()} AS [Create_User]
      ,{_UserInfo._date.dateTimeToStrSql()} AS [Create_Date]
      ,{_UserInfo._time.timeSpanToStrSql()} AS [Create_Time]
  FROM [Bond_Quantitative_Resource]
  where Report_Date = {dtStrSql}
  and Version = {versionSql}
  and Assessment_Sub_Kind = '{otherStr}'
)
INSERT INTO [Bond_Quantitative_Resource]
           ([Reference_Nbr]
           ,[Bond_Number]
           ,[Lots]
           ,[Segment_Name]
           ,[Version]
           ,[Report_Date]
           ,[Processing_Date]
           ,[Assessment_Sub_Kind]
           ,[Net_Debt]
           ,[Total_Asset]
           ,[CFO]
           ,[Int_Expense]
           ,[CET1_Capital_Ratio]
           ,[Lerverage_Ratio]
           ,[Liquiduty_Coverage_Ratio]
           ,[Total_Equity]
           ,[BS_TOT_ASSET]
           ,[RBC_Ratio]
           ,[MCCSR_Ratio]
           ,[Solvency_II_Ratio]
           ,[SHORT_AND_LONG_TERM_DEBT]
           ,[Cash_and_Cash_Equivalents]
           ,[Short_Term_Debt]
           ,[ISSUER_AREA]
           ,[IGS_INDEX]
           ,[IGS_INDEX_Increase]
           ,[External_Debt_Rate]
           ,[FC_Indexed_Debt_Rate]
           ,[CEIC_Value]
           ,[Short_term_Debt_Ratio]
           ,[GDP_Yearly_Avg]
           ,[Quantitative_CLO_1]
           ,[Quantitative_CLO_2]
           ,[Quantitative_CLO_2_Threshold]
           ,[Quantitative_CLO_3]
           ,[Quantitative_CLO_3_Threshold]
           ,[Quantitative_CLO_4]
           ,[Quantitative_CLO_4_Threshold]
           ,[Portfolio]
           ,[Origination_Date]
           ,[Portfolio_Name]
           ,[Assessment_Result_Version]
           ,[Assessor]
           ,[Assessment_date]
           ,[Auditor]
           ,[Audit_date]
           ,[Send_to_Auditor]
           ,[Send_Time]
           ,[Result_Version_Confirm]
           ,[Quantitative_Pass_Confirm]
           ,[Create_User]
           ,[Create_Date]
           ,[Create_Time]
)
select * from TEMPD63Other;
";
                #endregion

                try
                {
                    int i = db.Database.ExecuteSqlCommand(sql);
                    db.SaveChanges();
                    result.RETURN_FLAG = true;
                    result.DESCRIPTION = "執行量化評估需求成功";
                }
                catch (Exception ex)
                {
                    result.RETURN_FLAG = false;
                    result.DESCRIPTION = ex.exceptionMessage();
                }
            }
            return result;
        }
        #endregion

        #region 執行質化評估需求
        /// <summary>
        /// A41 data to D65
        /// </summary>
        /// <param name="reportDate"></param>
        /// <returns></returns>
        public MSGReturnModel TransferD65(string reportDate)
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            DateTime startDt = DateTime.Now;
            DateTime dt = DateTime.MinValue;
            int version = 0;
            if (!DateTime.TryParse(reportDate, out dt))
            {
                result.DESCRIPTION = Message_Type.parameter_Error.GetDescription();
                return result;
            }
            if (getD54Check(dt))
            {
                result.DESCRIPTION = Message_Type.already_Save_D54.GetDescription($"{dt.ToString("yyyyMMdd")}資料");
                return result;
            }
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                var D62 = db.Bond_Basic_Assessment.AsNoTracking().Where(x => x.Report_Date == dt);
                if (D62.Any())
                {
                    version = D62.Max(x => x.Version);

                    var _report = dt.ToString("yyyyMMdd");
                    var _cr = $"{_report}_V{version}";
                    var D66Files = db.Bond_Qualitative_Assessment_Result_File
                        .Where(x => x.Check_Reference.StartsWith(_cr)).ToList();

                    D66Files.ForEach(x =>
                    {
                        try
                        {
                            File.Delete(x.File_path);
                        }
                        catch
                        {

                        }
                    });

                    db.Bond_Qualitative_Assessment_Result_File.RemoveRange(D66Files);

                    var watchCount = D62.Where(x => x.Watch_IND == "Y" && x.Version == version).Count();
                    var D63 = db.Bond_Quantitative_Resource.AsNoTracking().
                        Where(x => x.Report_Date == dt && x.Version == version).ToList();
                    if (watchCount == 0)
                    {
                        string delSql = string.Empty;

                        delSql = $@"
DELETE Bond_Qualitative_Assessment
WHERE Report_Date = @Report_Date
AND Version = @Version ;

DELETE Bond_Qualitative_Assessment_Result
WHERE Report_Date = @Report_Date
AND Version = @Version ; ";

                        try
                        {
                            int i = db.Database.ExecuteSqlCommand(delSql,
                                new List<SqlParameter>()
                                {
                                  new SqlParameter("Report_Date", dt.ToString("yyyy-MM-dd")),
                                  new SqlParameter("Version", version.ToString())
                                }.ToArray());

                            db.SaveChanges();
                        }
                        catch
                        {

                        }

                        result.DESCRIPTION = "D62最大版本沒有符合觀察名單(需進行質化)的資料";
                        return result;
                    }
                    if (D63.Any(x => x.Result_Version_Confirm == null))
                    {
                        result.DESCRIPTION = "量化評估資料未全部完成複核,請查看複核狀況!";
                        return result;
                    }
                    if (!D63.Any(x => x.Result_Version_Confirm == x.Assessment_Result_Version &&
                                       x.Quantitative_Pass_Confirm == "N"))
                    {
                        result.DESCRIPTION = "量化評估資料全部通過,無須執行質化評估!";
                        return result;
                    }
                }
                else
                {
                    result.DESCRIPTION = Message_Type.not_Find_Any.GetDescription("D62", "無法執行!");
                    return result;
                }



                #region sql
                string sql = string.Empty;
                var dtStrSql = dt.dateTimeToStrSql();
                var versionSql = version.intToStrSql();
                sql = $@"
DELETE Bond_Qualitative_Assessment
WHERE Report_Date = {dtStrSql}
AND Version = {versionSql} ;

DELETE Bond_Qualitative_Assessment_Result
WHERE Report_Date = {dtStrSql}
AND Version = {versionSql} ;

WITH TEMP AS
(
    SELECT 
	Reference_Nbr AS Reference_Nbr,
	Version AS Version,
	Report_Date AS Report_Date,
	'質化衡量' AS Assessment_Kind,
	Assessment_Sub_Kind AS Assessment_Sub_Kind,
	Bond_Number AS Bond_Number,
	Lots AS Lots,
	Portfolio AS Portfolio,
	Origination_Date AS Origination_Date,
	Portfolio_Name AS Portfolio_Name
	FROM  Bond_Quantitative_Resource 
    where Report_Date = {dtStrSql}
    and Version = {versionSql} 
    and Result_Version_Confirm is not null
    and Result_Version_Confirm = Assessment_Result_Version
    and Quantitative_Pass_Confirm = 'N'
)
INSERT INTO
Bond_Qualitative_Assessment
(
Reference_Nbr,
Version,
Report_Date,
Assessment_Kind,
Assessment_Sub_Kind,
Bond_Number,
Lots,
Portfolio,
Origination_Date,
Portfolio_Name,
Assessment_Result_Version,
Create_User,
Create_Date,
Create_Time
)
SELECT 
Reference_Nbr,
Version,
Report_Date,
Assessment_Kind,
Assessment_Sub_Kind,
Bond_Number,
Lots,
Portfolio,
Origination_Date,
Portfolio_Name,
1,
{_UserInfo._user.stringToStrSql()},
{_UserInfo._date.dateTimeToStrSql()},
{_UserInfo._time.timeSpanToStrSql()}
FROM TEMP ";
                #endregion
                try
                {
                    int i = db.Database.ExecuteSqlCommand(sql);
                    db.SaveChanges();
                    result.RETURN_FLAG = true;
                    result.DESCRIPTION = "執行質化評估需求成功";
                }
                catch (Exception ex)
                {
                    result.RETURN_FLAG = false;
                    result.DESCRIPTION = ex.exceptionMessage();
                }
            }
            return result;
        }
        #endregion

        #region 量化評估推送
        /// <summary>
        /// D63 複核推送
        /// </summary>
        /// <param name="action">狀態 只有Edit</param>
        /// <param name="Auditor">覆核者</param>
        /// <param name="model"></param>
        /// <returns></returns>
        public MSGReturnModel SaveD63(string action, string Auditor, D63ViewModel model)
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            DateTime startDt = DateTime.Now;
            if (action == Action_Type.Edit.ToString())
            {
                int version = 1;
                Int32.TryParse(model.Version, out version);
                using (IFRS9DBEntities db = new IFRS9DBEntities())
                {
                    var D63s = db.Bond_Quantitative_Resource.Where(x => 
                                x.Reference_Nbr == model.Reference_Nbr &&
                                x.Version == version);
                    var D63 = D63s.FirstOrDefault(x => x.Assessment_Result_Version == 1);
                    var lastVersion = 0;
                    lastVersion = D63s.Max(x => x.Assessment_Result_Version) + 1;                          
                    if (D63 != null && D63.Result_Version_Confirm == null)
                    {
                        Bond_Quantitative_Resource D63Add =
                            D63.ModelConvert<Bond_Quantitative_Resource, Bond_Quantitative_Resource>();
                        List<Bond_Quantitative_Result> D64Datas = new List<Bond_Quantitative_Result>();
                        List<Bond_Assessment_Parm> D61s =
                            db.Bond_Assessment_Parm.AsNoTracking().Where(x =>
                                x.Assessment_Stage == QuantifyAssessmentStage && //Assessment_Stage == 2
                                x.Assessment_Kind == QuantifyAssessmentKind && // Assessment_Kind == 量化衡量
                                x.Assessment_Sub_Kind == D63.Assessment_Sub_Kind && // Assessment_Sub_Kind == D63.Assessment_Sub_Kind
                                x.IsActive == "Y" //生效
                            ).ToList();
                        if(D61s.Any())
                            //排序 把 Pass_Count_% 排到最後 要做 pass_value 相加用
                            D61s = D61s.OrderBy(x => x.Check_Item_Code.StartsWith(checkItemCode)).ToList();

                        if (D63.Assessment_Sub_Kind == AssessmentSubKind_Type.Industrial_Company.GetDescription())
                        {
                            D63Add.Net_Debt = TypeTransfer.stringToDoubleN(model.Net_Debt);
                            D63Add.Total_Asset = TypeTransfer.stringToDoubleN(model.Total_Asset);
                            D63Add.CFO = TypeTransfer.stringToDoubleN(model.CFO);
                            D63Add.Int_Expense = TypeTransfer.stringToDoubleN(model.Int_Expense);
                        }
                        else if (D63.Assessment_Sub_Kind == AssessmentSubKind_Type.Financial_Debts.GetDescription() ||
                                D63.Assessment_Sub_Kind == AssessmentSubKind_Type.Financial_Debt_Bond.GetDescription())
                        {
                            D63Add.CET1_Capital_Ratio = TypeTransfer.stringToDoubleN(model.CET1_Capital_Ratio);
                            D63Add.Lerverage_Ratio = TypeTransfer.stringToDoubleN(model.Lerverage_Ratio);
                            D63Add.Liquiduty_Coverage_Ratio = TypeTransfer.stringToDoubleN(model.Liquiduty_Coverage_Ratio);
                        }
                        else if (D63.Assessment_Sub_Kind == AssessmentSubKind_Type.Life_Or_Property_Insurance_Company.GetDescription())
                        {
                            D63Add.Total_Equity = TypeTransfer.stringToDoubleN(model.Total_Equity);
                            D63Add.BS_TOT_ASSET = TypeTransfer.stringToDoubleN(model.BS_TOT_ASSET);
                            D63Add.RBC_Ratio = TypeTransfer.stringToDoubleN(model.RBC_Ratio);
                            D63Add.MCCSR_Ratio = TypeTransfer.stringToDoubleN(model.MCCSR_Ratio);
                            D63Add.Solvency_II_Ratio = TypeTransfer.stringToDoubleN(model.Solvency_II_Ratio);
                            D63Add.SHORT_AND_LONG_TERM_DEBT = TypeTransfer.stringToDoubleN(model.SHORT_AND_LONG_TERM_DEBT);
                            D63Add.Cash_and_Cash_Equivalents = TypeTransfer.stringToDoubleN(model.Cash_and_Cash_Equivalents);
                            D63Add.Short_Term_Debt = TypeTransfer.stringToDoubleN(model.Short_Term_Debt);
                        }
                        else if (D63.Assessment_Sub_Kind == AssessmentSubKind_Type.CLO.GetDescription())
                        {
                            D63Add.Quantitative_CLO_1 = TypeTransfer.stringToDoubleN(model.Quantitative_CLO_1);
                            D63Add.Quantitative_CLO_2 = TypeTransfer.stringToDoubleN(model.Quantitative_CLO_2);
                            D63Add.Quantitative_CLO_2_Threshold = TypeTransfer.stringToDoubleN(model.Quantitative_CLO_2_Threshold);
                            D63Add.Quantitative_CLO_3 = TypeTransfer.stringToDoubleN(model.Quantitative_CLO_3);
                            D63Add.Quantitative_CLO_3_Threshold = TypeTransfer.stringToDoubleN(model.Quantitative_CLO_3_Threshold);
                            D63Add.Quantitative_CLO_4 = TypeTransfer.stringToDoubleN(model.Quantitative_CLO_4);
                            D63Add.Quantitative_CLO_4_Threshold = TypeTransfer.stringToDoubleN(model.Quantitative_CLO_4_Threshold);
                        }
                        else if (D63.Assessment_Sub_Kind == AssessmentSubKind_Type.Sovereign_Government_Debt.GetDescription())
                        {
                            D63Add.IGS_INDEX = TypeTransfer.stringToDoubleN(model.IGS_INDEX);
                            D63Add.IGS_INDEX_Increase = TypeTransfer.stringToDoubleN(model.IGS_INDEX_Increase);
                            D63Add.External_Debt_Rate = TypeTransfer.stringToDoubleN(model.External_Debt_Rate);
                            D63Add.FC_Indexed_Debt_Rate = TypeTransfer.stringToDoubleN(model.FC_Indexed_Debt_Rate);
                            D63Add.CEIC_Value = TypeTransfer.stringToDoubleN(model.CEIC_Value);
                            D63Add.Short_term_Debt_Ratio = TypeTransfer.stringToDoubleN(model.Short_term_Debt_Ratio);
                            D63Add.GDP_Yearly_Avg = TypeTransfer.stringToDoubleN(model.GDP_Yearly_Avg);
                        }
                        //參數初始化
                        D63Add.Auditor_Return = null;
                        D63Add.Send_Time = null;
                        D63Add.Auditor_Reply = null;
                        D63Add.Send_to_Auditor = null;
                        D63Add.LastUpdate_User = null;
                        D63Add.LastUpdate_Date = null;
                        D63Add.LastUpdate_Time = null;
                        
                        double totalPassValue = 0d;
                        List<string> passCode = new List<string>(); //比對X53,X54 使用,獲取通過的Check_Item_Code
                        List<string> expassCode = new List<string>() {"X53","X54"};
                        D61s.ForEach(x =>
                        {
                            //量化指標值
                            var _Check_Item_Value = D64ItemValue(x.Check_Item_Code, D63Add);
                            Bond_Quantitative_Result D64 = new Bond_Quantitative_Result();
                            D64.Reference_Nbr = D63.Reference_Nbr;
                            D64.Report_Date = D63.Report_Date;
                            D64.Version = D63.Version;
                            D64.Assessment_Kind = QuantifyAssessmentKind;
                            D64.Assessment_Sub_Kind = D63.Assessment_Sub_Kind;
                            D64.Assessment_Stage = "2";
                            D64.Check_Item_Code = x.Check_Item_Code;
                            D64.Check_Item = x.Check_Item;
                            D64.Check_Item_Memo = x.Check_Item_Memo;
                            D64.Check_Reference = string.Format("{0}_V{1}_{2}_ARV{3}_{4}",
                            D64.Report_Date.ToString("yyyyMMdd"), 
                            D64.Version.ToString(),
                            D64.Reference_Nbr,
                            lastVersion.ToString(), 
                            D64.Check_Item_Code);
                            D64.Assessment_Result_Version = lastVersion;                          
                            // Pass_Count_X5 特殊例外  X53 & X54 都符合時
                            if (D64.Check_Item_Code.StartsWith(checkItemCode) &&
                            expassCode.Intersect(passCode).Count() == expassCode.Count)
                            {
                                // X53 & X54 都符合時 只算一次(各算0.5 => 所以要減一)
                                var changes = D64Datas.Where(z => expassCode.Contains(z.Check_Item_Code)).ToList();
                                changes.ForEach(w => { w.Pass_Value = 0.5; });
                                totalPassValue = totalPassValue - 1;
                            }
                            D64.Check_Item_Value = D64.Check_Item_Code.StartsWith(checkItemCode) ?
                            totalPassValue : _Check_Item_Value;
                            D64.Pass_Value = D64.Check_Item_Code.StartsWith(checkItemCode) ? //Pass_Count_% 為 total
                            D64PassValue(x,totalPassValue ,D63Add):
                            _Check_Item_Value.HasValue ? D64PassValue(x,_Check_Item_Value.Value, D63Add) : 0d;                               
                            D64.Check_Symbol = x.Check_Symbol;
                            D64.Threshold = D64Threshold(x, D63Add);
                            D64.Bond_Number = D63.Bond_Number;
                            D64.Lots = D63.Lots;
                            D64.Portfolio = D63.Portfolio;
                            D64.Origination_Date = D63.Origination_Date;
                            D64.Portfolio_Name = D63.Portfolio_Name;
                            D64.Quantitative_Pass = D64QuantitativePass(x, D64.Check_Item_Value.HasValue ? D64.Check_Item_Value.Value : 0d, D64.Pass_Value.HasValue ? D64.Pass_Value.Value : 0d, D63Add.CFO);
                            if (D64.Check_Item_Code.StartsWith(checkItemCode))
                            {
                                D63Add.Quantitative_Pass_Confirm = D64.Quantitative_Pass;
                            }
                            D64.Create_User = _UserInfo._user;
                            D64.Create_Date = _UserInfo._date;
                            D64.Create_Time = _UserInfo._time;
                            D64Datas.Add(D64);
                            if (!D64.Check_Item_Code.StartsWith(checkItemCode))
                                totalPassValue += D64.Pass_Value.Value;
                            if (D64.Pass_Value != null && D64.Pass_Value != 0)
                                passCode.Add(D64.Check_Item_Code);                                
                        });
                        D63Add.Assessment_Result_Version = lastVersion;
                        D63Add.Processing_Date = startDt;
                        D63Add.Assessor = _UserInfo._user;
                        D63Add.Assessment_date = _UserInfo._date;
                        D63Add.Auditor = Auditor;
                        D63Add.Create_User = _UserInfo._user;
                        D63Add.Create_Date = _UserInfo._date;
                        D63Add.Create_Time = _UserInfo._time;                       
                        db.Bond_Quantitative_Resource.Add(D63Add);
                        db.Bond_Quantitative_Result.AddRange(D64Datas);
                        try
                        {
                            db.SaveChanges();
                            result.RETURN_FLAG = true;
                            result.DESCRIPTION = Message_Type.push_Assessor_Success.GetDescription();
                        }
                        catch (DbUpdateException ex)
                        {
                            result.RETURN_FLAG = false;
                            result.DESCRIPTION = ex.exceptionMessage();
                        }
                    }
                    else
                    {
                        result.DESCRIPTION = Message_Type.already_Change.GetDescription();
                    }
                }
            }
            return result;
        }
        #endregion

        #region 質化評估推送
        /// <summary>
        /// D65 複核推送
        /// </summary>
        /// <param name="Auditor"></param>
        /// <param name="D65Model">存D65 info ref_Nbr </param>
        /// <param name="D66Models">存D66 info 前端只送 pass_value & Check_Item_Code</param>
        /// <returns></returns>
        public MSGReturnModel SaveD65(string Auditor, D65ViewModel D65Model, List<D66ViewModel> D66Models)
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            if (!D66Models.Any())
            {
                result.DESCRIPTION = Message_Type.parameter_Error.GetDescription();
                return result;
            }
            int version = 1;
            Int32.TryParse(D65Model.Version, out version);
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                var D65s = db.Bond_Qualitative_Assessment.AsNoTracking()
                    .Where(x => x.Reference_Nbr == D65Model.Reference_Nbr &&
                                x.Version == version).AsEnumerable();
                var D65 = D65s.FirstOrDefault(x => x.Assessment_Result_Version == 1);
                var lastVersion = 0;
                lastVersion = D65s.Max(x => x.Assessment_Result_Version) + 1;
                if (D65 != null && D65.Result_Version_Confirm == null && D65.Extra_Case != "D")
                {
                    Bond_Qualitative_Assessment D65Add =
                        D65.ModelConvert<Bond_Qualitative_Assessment, Bond_Qualitative_Assessment>();
                    List<Bond_Qualitative_Assessment_Result> D66Datas = new List<Bond_Qualitative_Assessment_Result>();
                    var Assessment_Stage = D66Models.First().Assessment_Stage;
                    List<Bond_Assessment_Parm> D61s =
                        db.Bond_Assessment_Parm.AsNoTracking().Where(x =>
                            x.Assessment_Stage == Assessment_Stage &&
                            x.Assessment_Kind == QualitativeAssessmentKind &&
                            x.Assessment_Sub_Kind == D65Model.Assessment_Sub_Kind &&
                            x.IsActive == "Y" //生效
                        ).ToList();

                    if (D61s.Any())
                        //排序 把 Pass_Count_% 排到最後 要做 pass_value 相加用
                        D61s = D61s.OrderBy(x => x.Check_Item_Code.StartsWith(checkItemCode)).ToList();

                    //參數初始化
                    D65Add.Auditor_Return = null;
                    D65Add.Send_Time = null;
                    D65Add.Auditor_Reply = null;
                    D65Add.Send_to_Auditor = null;
                    D65Add.LastUpdate_User = null;
                    D65Add.LastUpdate_Date = null;
                    D65Add.LastUpdate_Time = null;

                    double totalPassValue = 0d;
                    D61s.ForEach(x =>
                    {
                        var D66Model = D66Models.FirstOrDefault(z => z.Check_Item_Code == x.Check_Item_Code);
                        Bond_Qualitative_Assessment_Result D66 = new Bond_Qualitative_Assessment_Result();
                        D66.Bond_Number = D65.Bond_Number;
                        D66.Reference_Nbr = D65.Reference_Nbr;
                        D66.Report_Date = D65.Report_Date;
                        D66.Version = D65.Version;
                        D66.Assessment_Kind = QualitativeAssessmentKind;
                        D66.Assessment_Sub_Kind = x.Assessment_Sub_Kind;
                        D66.Assessment_Stage = x.Assessment_Stage;
                        D66.Check_Item_Code = x.Check_Item_Code;
                        D66.Check_Item = x.Check_Item;
                        D66.Check_Item_Memo = x.Check_Item_Memo;
                        D66.Check_Reference = string.Format("{0}_V{1}_{2}_ARV{3}_{4}",
                            D66.Report_Date.ToString("yyyyMMdd"), 
                            D66.Version.ToString(),
                            D66.Reference_Nbr,
                            lastVersion.ToString(), 
                            D66.Check_Item_Code);
                        D66.Assessment_Result_Version = lastVersion;
                        D66.Check_Symbol = x.Check_Symbol;
                        D66.Threshold = x.Threshold;
                        D66.Pass_Value = D66.Check_Item_Code.StartsWith(checkItemCode) ? totalPassValue :
                        (D66Model.Pass_Value == "1" ? 1 : 0);
                        D66.Qualitative_Pass_Stage2 = D66QuantitativePass(D66.Check_Item_Code, D66.Pass_Value.Value, D66.Assessment_Stage, "2", x.Threshold, D66.Check_Symbol);
                        D66.Qualitative_Pass_Stage3 = D66QuantitativePass(D66.Check_Item_Code, D66.Pass_Value.Value, D66.Assessment_Stage, "3", x.Threshold, D66.Check_Symbol);
                        if (D66.Check_Item_Code.StartsWith(checkItemCode))
                        {
                            D65Add.Quantitative_Pass_Confirm = D66.Qualitative_Pass_Stage2;
                        }
                        if (D66.Check_Item_Code.StartsWith("Z"))
                        {
                            D65Add.Quantitative_Pass_Confirm = D66.Qualitative_Pass_Stage3;
                        }
                        D66.Create_User = _UserInfo._user;
                        D66.Create_Date = _UserInfo._date;
                        D66.Create_Time = _UserInfo._time;
                        D66Datas.Add(D66);
                        if (!D66.Check_Item_Code.StartsWith(checkItemCode))
                            totalPassValue += D66.Pass_Value.Value;
                    });
                    D65Add.Assessment_Result_Version = lastVersion;
                    D65Add.Assessor = _UserInfo._user;
                    D65Add.Assessment_date = _UserInfo._date;
                    D65Add.Auditor = Auditor;
                    D65Add.Create_User = _UserInfo._user;
                    D65Add.Create_Date = _UserInfo._date;
                    D65Add.Create_Time = _UserInfo._time;
                    db.Bond_Qualitative_Assessment.Add(D65Add);
                    db.Bond_Qualitative_Assessment_Result.AddRange(D66Datas);
                    try
                    {
                        db.SaveChanges();
                        result.RETURN_FLAG = true;
                        result.DESCRIPTION = Message_Type.push_Assessor_Success.GetDescription();
                    }
                    catch (DbUpdateException ex)
                    {
                        result.RETURN_FLAG = false;
                        result.DESCRIPTION = ex.exceptionMessage();
                    }
                }
            }
            return result;
        }
        #endregion

        #region save QualitativeFile
        /// <summary>
        /// save QualitativeFile
        /// </summary>
        /// <param name="Reference_Nbr"></param>
        /// <param name="Check_Reference"></param>
        /// <param name="FileName"></param>
        /// <param name="type">D64.D66</param>
        /// <param name="File_path">檔案位置</param>
        /// <returns></returns>
        public MSGReturnModel SaveQualitativeFile(string Check_Reference, string FileName,Table_Type type,string File_path)
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            if (!(type  == Table_Type.D64
                || type  == Table_Type.D66))
            {
                result.DESCRIPTION = Message_Type.parameter_Error.GetDescription();
                return result;
            }
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                if (type  == Table_Type.D64)
                {
                    var D64 = db.Bond_Quantitative_Result.AsNoTracking()
                        .Where(x => x.Check_Reference == Check_Reference).FirstOrDefault() ?? new Bond_Quantitative_Result();
                    var D63 = db.Bond_Quantitative_Resource.AsNoTracking().FirstOrDefault(x => x.Reference_Nbr == D64.Reference_Nbr);
                    var D64File = db.Bond_Quantitative_Result_File
                         .FirstOrDefault(x => x.Check_Reference == Check_Reference &&
                         x.File_Name == FileName);
                    if (D63 == null || (D64File != null && D64File.Status == "Y"))
                    {
                        result.DESCRIPTION = Message_Type.already_Save.GetDescription();
                        return result;
                    }
                    D63.LastUpdate_User = _UserInfo._user;
                    D63.LastUpdate_Date = _UserInfo._date;
                    D63.LastUpdate_Time = _UserInfo._time;
                    if (D64File != null && D64File.Status == "N")
                    {
                        D64File.Create_User = _UserInfo._user;
                        D64File.Create_Date = _UserInfo._date;
                        D64File.Create_Time = _UserInfo._time;
                        D64File.LastUpdate_User = null;
                        D64File.LastUpdate_Date = null;
                        D64File.LastUpdate_Time = null;
                        D64File.Status = "Y";
                    }
                    else
                    {
                        db.Bond_Quantitative_Result_File.Add(
                        new Bond_Quantitative_Result_File()
                        {
                            Check_Reference = Check_Reference,
                            File_Name = FileName,
                            Create_User = _UserInfo._user,
                            Create_Date = _UserInfo._date,
                            Create_Time = _UserInfo._time,
                            File_path = File_path,
                            Status = "Y"
                        });
                    }
                }
                if (type  == Table_Type.D66)
                {
                    var D66 = db.Bond_Qualitative_Assessment_Result.AsNoTracking()
                            .Where(x => x.Check_Reference == Check_Reference).FirstOrDefault() ?? new Bond_Qualitative_Assessment_Result();
                    var D65 = db.Bond_Qualitative_Assessment.AsNoTracking().FirstOrDefault(x => x.Reference_Nbr == D66.Reference_Nbr);
                    var D66File = db.Bond_Qualitative_Assessment_Result_File
                          .FirstOrDefault(x => x.Check_Reference == Check_Reference &&
                          x.File_Name == FileName);
                    if (D65 == null || (D66File != null && D66File.Status == "Y"))
                    {
                        result.DESCRIPTION = Message_Type.already_Save.GetDescription();
                        return result;
                    }
                    D65.LastUpdate_User = _UserInfo._user;
                    D65.LastUpdate_Time = _UserInfo._time;
                    D65.LastUpdate_Date = _UserInfo._date;
                    if (D66File != null && D66File.Status == "N")
                    {
                        D66File.Create_User = _UserInfo._user;
                        D66File.Create_Date = _UserInfo._date;
                        D66File.Create_Time = _UserInfo._time;
                        D66File.LastUpdate_User = null;
                        D66File.LastUpdate_Date = null;
                        D66File.LastUpdate_Time = null;
                        D66File.Status = "Y";
                    }
                    else
                    {
                        db.Bond_Qualitative_Assessment_Result_File.Add(
                        new Bond_Qualitative_Assessment_Result_File()
                        {
                        Check_Reference = Check_Reference,
                        File_Name = FileName,
                        Create_User = _UserInfo._user,
                        Create_Date = _UserInfo._date,
                        Create_Time = _UserInfo._time,
                        File_path = File_path,
                        Status = "Y"
                        });
                    }

                }
                try
                {
                    db.SaveChanges();
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
        #endregion

        #region 量化化評估檔案動作
        /// <summary>
        /// D63 Event
        /// </summary>
        /// <param name="Reference_Nbr"></param>
        /// <param name="Version"></param>
        /// <param name="Assessment_Result_Version"></param>
        /// <param name="status"></param>
        /// <returns></returns>
        public MSGReturnModel UpdateD63(
            string Reference_Nbr,
            int Assessment_Result_Version,
            Evaluation_Status_Type status)
        {
            MSGReturnModel result = new MSGReturnModel();
            result = D6ReviewMethod(
                Reference_Nbr,
                Assessment_Result_Version,
                Table_Type.D63,
                status);

            if (result.RETURN_FLAG)
            {
                using (IFRS9DBEntities db = new IFRS9DBEntities())
                {
                    var _D63 = db.Bond_Quantitative_Resource.AsNoTracking()
                        .FirstOrDefault(x => x.Reference_Nbr == Reference_Nbr);
                    if (_D63 != null)
                    {
                        var _D63s = db.Bond_Quantitative_Resource.AsNoTracking()
                            .Where(x =>
                            x.Report_Date == _D63.Report_Date &&
                            x.Version == _D63.Version).ToList();
                        //條件1 所有D63都已經複核確認  ,條件2 
                        if (!_D63s.Any(x => x.Result_Version_Confirm == null) &&
                            _D63s.Any(x => x.Assessment_Result_Version == x.Result_Version_Confirm &&
                                             x.Quantitative_Pass_Confirm != "Y")
                            )
                        {
                            var _reportDate = $" 基準日:{ _D63.Report_Date.ToString("yyyy/MM/dd")}";
                            var log = string.Empty;
                            log += $@"寄信時間:{DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss")}";
                            log += _reportDate;
                            var sms = new SendMail.SendMailSelf();
                            //sms.smtpServer = "msa.hinet.net";
                            sms.smtpPort = 25;
                            sms.smtpServer = Properties.Settings.Default["smtpServer"]?.ToString(); 
                            sms.mailAccount = Properties.Settings.Default["mailAccount"]?.ToString();
                            //sms.mailPwd = Properties.Settings.Default["mailPwd"]?.ToString();
                            //var mail = db.Notice_Info.AsNoTracking()
                            //    .FirstOrDefault(x => x.IsActive == "Y" && x.Notice_Name == "量化作業已完成");
                           var _Notice_ID = TypeTransfer.stringToInt(Properties.Settings.Default["D63MailID"].ToString());
                            var mail = db.Notice_Info.AsNoTracking()
                               .FirstOrDefault(x => x.Notice_ID == _Notice_ID);
                            if (mail != null)
                            {
                                log += " 收件者:";
                                var mailtos = new List<Tuple<string, string>>();
                                bool flag = false;
                                foreach (var item in mail.Mail_Info.Where(x => x.IsActive == "Y"))
                                {
                                    flag = true;
                                    mailtos.Add(new Tuple<string, string>(item.Recipient_mail, item.Recipient_Name));
                                    log += $@"{item.Recipient_mail}({item.Recipient_Name}),";
                                }
                                if (flag)
                                    log = log.Substring(0, log.Length - 1);
                                var _msg = sms.Mail_Send(
                                //new Tuple<string, string>("glsisys.life@fbt.com", "測試帳號-glsisys"),
                                new Tuple<string, string>(sms.mailAccount,"IFRS9系統帳號"),
                                mailtos,
                                null,
                                $"{mail.Mail_Title}  {_reportDate}",
                                $"{mail.Mail_Msg}  {_reportDate}");
                                if (_msg.IsNullOrWhiteSpace())                              
                                    log += " 結果:傳送郵件成功。";
                                else
                                    log += " 結果:傳送郵件失敗。";                                                            
                                Extension.NlogSet($"寄信 : {_msg}");
                            }
                            else
                            {
                                log += " 結果:無生效寄信設定";
                            }
                            result.REASON_CODE = log;
                        }
                    }
                }
            }
            return result;
        }
        #endregion

        #region 紀錄寄信到 txt 檔案

        private void saveMailtxt()
        {

        }

        #endregion

        #region 質化評估檔案動作
        /// <summary>
        /// D65 Event
        /// </summary>
        /// <param name="Reference_Nbr"></param>
        /// <param name="Version"></param>
        /// <param name="Assessment_Result_Version"></param>
        /// <param name="status"></param>
        /// <returns></returns>
        public MSGReturnModel UpdateD65(
            string Reference_Nbr,
            int Assessment_Result_Version,
            Evaluation_Status_Type status)
        {
            MSGReturnModel result = new MSGReturnModel();
            result = D6ReviewMethod(
                Reference_Nbr,
                Assessment_Result_Version,
                Table_Type.D65, 
                status);
            return result;
        }
        #endregion

        /// <summary>
        /// 刪除D64orD66附件檔案
        /// </summary>
        /// <param name="type"></param>
        /// <param name="Check_Reference"></param>
        /// <param name="fileName"></param>
        /// <returns></returns>
        public MSGReturnModel DelQuantifyAndQualitativeFile(string type,string Check_Reference,string fileName)
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                Bond_Quantitative_Result_File D64 = null;
                if (type == "D64")
                {
                    D64 = db.Bond_Quantitative_Result_File
                       .FirstOrDefault(x => x.Check_Reference == Check_Reference &&
                       x.File_Name == fileName);
                    if (D64 != null)
                    {
                        D64.Status = "N";
                        D64.LastUpdate_User = _UserInfo._user;
                        D64.LastUpdate_Date = _UserInfo._date;
                        D64.LastUpdate_Time = _UserInfo._time;
                    }
                }
                Bond_Qualitative_Assessment_Result_File D66 = null;
                if (type == "D66")
                {
                    D66 = db.Bond_Qualitative_Assessment_Result_File
                        .FirstOrDefault(x => x.Check_Reference == Check_Reference &&
                        x.File_Name == fileName);
                    if (D66 != null)
                    {
                        D66.Status = "N";
                        D66.LastUpdate_User = _UserInfo._user;
                        D66.LastUpdate_Date = _UserInfo._date;
                        D66.LastUpdate_Time = _UserInfo._time;
                    }              
                }
                if (D64 != null || D66 != null)
                {
                    try
                    {
                        if(D64 != null)
                            System.IO.File.Delete(D64.File_path);
                        if (D66 != null)
                            System.IO.File.Delete(D66.File_path);
                        db.SaveChanges();
                        result.RETURN_FLAG = true;
                        result.DESCRIPTION = Message_Type.delete_Success.GetDescription();
                    }
                    catch (Exception ex)
                    {
                        result.DESCRIPTION = ex.exceptionMessage();
                    }
                }
                else
                {
                    result.DESCRIPTION = Message_Type.already_Change.GetDescription();
                }
            }
            return result;
        }

        #region 質化 save Assessment_Result , Memo
        /// <summary>
        /// save Assessment_Result , Memo
        /// </summary>
        /// <param name="value"></param>
        /// <param name="ReportDate"></param>
        /// <param name="Reference_Nbr"></param>
        /// <param name="Version"></param>
        /// <param name="Assessment_Result_Version"></param>
        /// <param name="Check_Item_Code"></param>
        /// <param name="type"></param>
        /// <param name="table"></param>
        /// <returns></returns>
        public MSGReturnModel SaveTextArea(
            string value,
            DateTime ReportDate,
            string Reference_Nbr,
            int Version,
            int Assessment_Result_Version,
            string Check_Item_Code,
            string type,
            Table_Type table)
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            if (!( table == Table_Type.D64
                || table == Table_Type.D66
                || table == Table_Type.D63
                || table == Table_Type.D65
                ))
            {
                result.DESCRIPTION = Message_Type.parameter_Error.GetDescription();
                return result;
            }
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                Bond_Quantitative_Resource D63 = null;
                Bond_Quantitative_Result D64 = null;
                Bond_Qualitative_Assessment D65 = null;
                Bond_Qualitative_Assessment_Result D66 = null;

                switch (table)
                {
                    case Table_Type.D63:
                        D63 = db.Bond_Quantitative_Resource
                            .FirstOrDefault(x => x.Reference_Nbr == Reference_Nbr &&
                            x.Report_Date == ReportDate &&
                            x.Version == Version &&
                            x.Assessment_Result_Version == Assessment_Result_Version);
                        break;
                    case Table_Type.D64:
                        D64 = db.Bond_Quantitative_Result
                            .FirstOrDefault(x => x.Reference_Nbr == Reference_Nbr &&
                            x.Report_Date == ReportDate &&
                            x.Version == Version &&
                            x.Assessment_Result_Version == Assessment_Result_Version &&
                            x.Check_Item_Code == Check_Item_Code);
                        break;
                    case Table_Type.D65:
                        D65 = db.Bond_Qualitative_Assessment
                            .FirstOrDefault(x => x.Reference_Nbr == Reference_Nbr &&
                            x.Report_Date == ReportDate &&
                            x.Version == Version &&
                            x.Assessment_Result_Version == Assessment_Result_Version);
                        break;
                    case Table_Type.D66:
                        D66 = db.Bond_Qualitative_Assessment_Result
                            .FirstOrDefault(x => x.Reference_Nbr == Reference_Nbr &&
                            x.Report_Date == ReportDate &&
                            x.Version == Version &&
                            x.Assessment_Result_Version == Assessment_Result_Version &&
                            x.Check_Item_Code == Check_Item_Code);
                        break;
                }
                if (D63 != null || D64 !=null || D65 != null || D66 != null)
                {
                    switch (table)
                    {
                        case Table_Type.D64:
                            if (type == "Assessment_Result")
                                D64.Assessment_Result = value;
                            if (type == "Memo")
                                D64.Memo = value;
                            D64.LastUpdate_User = _UserInfo._user;
                            D64.LastUpdate_Date = _UserInfo._date;
                            D64.LastUpdate_Time = _UserInfo._time;
                            break;
                        case Table_Type.D66:
                            if (type == "Assessment_Result")
                                D66.Assessment_Result = value;
                            if (type == "Memo")
                                D66.Memo = value;
                            D66.LastUpdate_User = _UserInfo._user;
                            D66.LastUpdate_Date = _UserInfo._date;
                            D66.LastUpdate_Time = _UserInfo._time;
                            break;
                        case Table_Type.D63:
                            if (type == "Auditor_Reply")
                                D63.Auditor_Reply = value;
                            D63.LastUpdate_User = _UserInfo._user;
                            D63.LastUpdate_Date = _UserInfo._date;
                            D63.LastUpdate_Time = _UserInfo._time;
                            break;
                        case Table_Type.D65:
                            if (type == "Auditor_Reply")
                                D65.Auditor_Reply = value;
                            D65.LastUpdate_User = _UserInfo._user;
                            D65.LastUpdate_Date = _UserInfo._date;
                            D65.LastUpdate_Time = _UserInfo._time;
                            break;
                    }
                    try
                    {
                        db.SaveChanges();
                        result.RETURN_FLAG = true;
                        result.DESCRIPTION = Message_Type.save_Success.GetDescription();
                    }
                    catch (DbUpdateException ex)
                    {
                        result.DESCRIPTION = ex.exceptionMessage();
                    }
                }
                else
                {
                    result.DESCRIPTION = Message_Type.not_Find_Any.GetDescription();
                }
            }
            return result;
        }
        #endregion

        public MSGReturnModel AutoInsertD65ExtraCase(List<Bond_Account_Info>A41Datas,string reportdate)
        {
            MSGReturnModel result = new MSGReturnModel();
            int _A41Ver = A41Datas.Max(x => x.Version).Value ;
            DateTime _ReportDate = new DateTime();
            DateTime.TryParse(reportdate, out _ReportDate);
            result.RETURN_FLAG = false;
            result.DESCRIPTION = "一鍵新增作業有誤";
            int num_insert = 0;
            int num_repeated = 0;
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                var _D62 = db.Bond_Basic_Assessment.Where(x => x.Report_Date == _ReportDate);               
                if(_D62.Any()&&_A41Ver==_D62.Max(x=>x.Version))
                {

                    foreach (var item in A41Datas)
                    {
                        if (db.Bond_Qualitative_Assessment.AsNoTracking().Any(
                            x => x.Reference_Nbr == item.Reference_Nbr &&
                            x.Report_Date == item.Report_Date &&
                            x.Bond_Number == item.Bond_Number &&
                            x.Lots == item.Lots &&
                            x.Portfolio_Name == item.Portfolio_Name
                            ))
                        {
                            num_repeated++;
                            continue;
                        }
                        else
                        {
                            db.Bond_Qualitative_Assessment.Add(
                                new Bond_Qualitative_Assessment()
                                {
                                    Reference_Nbr = item.Reference_Nbr,
                                    Version = item.Version.Value,
                                    Report_Date = item.Report_Date.Value,
                                    Assessment_Sub_Kind = item.Assessment_Sub_Kind,
                                    Bond_Number = item.Bond_Number,
                                    Lots = item.Lots,
                                    Portfolio = item.Portfolio,
                                    Origination_Date = item.Origination_Date,
                                    Portfolio_Name = item.Portfolio_Name,
                                    Assessment_Result_Version = 1,
                                    Assessment_Kind = "質化衡量",
                                    Extra_Case = "Y",
                                    Create_User = _UserInfo._user,
                                    Create_Date = _UserInfo._date,
                                    Create_Time = _UserInfo._time
                                }
                                );
                            num_insert++;
                        }

                    }
                    try
                    {
                        db.SaveChanges();
                        result.RETURN_FLAG = true;
                        result.DESCRIPTION = Message_Type.save_Success.GetDescription(null, $"重複個案{num_repeated}筆，新增個案{num_insert}筆，請自查詢頁面查詢內容。");
                    }
                    catch (Exception ex)
                    {
                        result.DESCRIPTION = ex.exceptionMessage();
                    }
                }
                else
                {
                    result.DESCRIPTION = "A41最大版本與基本要件評估最大版本不符";
                }
            }


            return result;
        }


        /// <summary>
        /// 新增D65 特殊案例
        /// </summary>
        /// <param name="reportDate"></param>
        /// <param name="version"></param>
        /// <param name="bondNumber"></param>
        /// <param name="Lots"></param>
        /// <param name="portfolio_Name"></param>
        /// <returns></returns>
        public MSGReturnModel InsertD65ByExtraCase(string reportDate, int version, string bondNumber, string Lots, string portfolio_Name)
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            result.DESCRIPTION = "A41找不到相關資料";
            DateTime _reportDate = DateTime.Parse(reportDate);
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                if (db.Bond_Qualitative_Assessment.AsNoTracking().Any(
                    x => x.Report_Date == _reportDate &&
                    x.Version == version &&
                    x.Bond_Number == bondNumber &&
                    x.Lots == Lots &&
                    x.Portfolio_Name == portfolio_Name
                    ))
                {
                    result.DESCRIPTION = "質化評估有相同的資料";
                    return result;
                }

                var _first = db.Bond_Account_Info.AsNoTracking()
                    .FirstOrDefault(x =>
                    x.Report_Date == _reportDate &&
                    x.Version == version &&
                    x.Bond_Number == bondNumber &&
                    x.Lots == Lots &&
                    x.Portfolio_Name == portfolio_Name);
                if (_first != null)
                {
                    db.Bond_Qualitative_Assessment.Add(
                        new Bond_Qualitative_Assessment() {
                            Reference_Nbr = _first.Reference_Nbr,
                            Version = _first.Version.Value,
                            Report_Date = _first.Report_Date.Value,
                            Assessment_Sub_Kind = _first.Assessment_Sub_Kind,
                            Bond_Number = _first.Bond_Number,
                            Lots = _first.Lots,
                            Portfolio = _first.Portfolio,
                            Origination_Date = _first.Origination_Date,
                            Portfolio_Name = _first.Portfolio_Name,
                            Assessment_Result_Version = 1,
                            Assessment_Kind = "質化衡量",
                            Extra_Case = "Y",
                            Create_User = _UserInfo._user,
                            Create_Date = _UserInfo._date,
                            Create_Time = _UserInfo._time
                        });
                    try
                    {
                        db.SaveChanges();
                        result.RETURN_FLAG = true;
                        result.DESCRIPTION = Message_Type.save_Success.GetDescription(null,"請自查詢頁面查詢");
                    }
                    catch (Exception ex)
                    {
                        result.DESCRIPTION = ex.exceptionMessage();
                    }
                }
            }
            return result;
        }

        /// <summary>
        /// 刪除 D65特殊案例
        /// </summary>
        /// <param name="referenceNbr"></param>
        /// <param name="version"></param>
        /// <returns></returns>
        public MSGReturnModel DeleteD65ByExtraCase(string referenceNbr, int version)
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            result.DESCRIPTION = Message_Type.delete_Fail.GetDescription();

            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                var datas = db.Bond_Qualitative_Assessment
                     .Where(x => x.Reference_Nbr == referenceNbr &&
                     x.Version == version);
                if (datas.Any())
                {
                    foreach (var data in datas)
                    {
                        db.Bond_Qualitative_Assessment.Remove(data); //190621 PGE需求發現BUG所作修改
                    }
                    try
                    {
                        db.SaveChanges();
                        result.RETURN_FLAG = true;
                        result.DESCRIPTION = Message_Type.delete_Success.GetDescription();
                    }
                    catch (Exception ex)
                    {
                        result.DESCRIPTION = ex.exceptionMessage();
                    }
                }
            }
            return result;
        }

        /// <summary>
        /// 減損報表資料刪除作業
        /// </summary>
        /// <param name="reportDate"></param>
        /// <param name="flowId"></param>
        /// <param name="msg"></param>
        /// <returns></returns>
        public MSGReturnModel ReEL(string reportDate, string flowId, string msg)
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            result.DESCRIPTION = Message_Type.not_Find_Any.GetDescription();
            string _reportDate = string.Empty;
            DateTime dt = DateTime.MinValue;
            if (DateTime.TryParse(reportDate, out dt))
            {
                _reportDate = dt.ToString("yyyy-MM-dd");
            }
            if (_reportDate.IsNullOrWhiteSpace() || flowId.IsNullOrWhiteSpace() || msg.IsNullOrWhiteSpace())
            {
                result.DESCRIPTION = Message_Type.parameter_Error.GetDescription();
                return result;
            }
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                var ver = db.IFRS9_EL.Where(x=>x.Report_Date==_reportDate).Max(x => x.Version).ToString();
                string sql = string.Empty;
                sql = $@"
                delete IFRS9_EL
                where Report_Date = {_reportDate.stringToStrSql()}
                and FLOWID = {flowId.stringToStrSql()} ;
                
                delete IFRS9_Bond_Report
                where Report_Date =  {_reportDate.stringToStrSql()}
                and FLOWID = {flowId.stringToStrSql()} ;
 
                delete Bond_Account_AssessmentCheck
                where Report_Date={_reportDate.stringToStrSql()}
                and Version={Convert.ToInt16(ver)};"; //190527 PGEv需求新增刪除C10 鎖定版本內容


                try
                {
                    using (var dbContextTransaction = db.Database.BeginTransaction())
                    {
                        int del = db.Database.ExecuteSqlCommand(sql);
                        if (del != 0)
                        {
                            db.IFRS9_RE_EL.Add(new IFRS9_RE_EL()
                            {
                                Report_Date = dt,
                                FLOWID = flowId,
                                Message = msg,
                                Create_User = _UserInfo._user,
                                Create_Date = DateTime.Now.Date,
                                Create_Time = DateTime.Now.TimeOfDay
                            });
                            db.SaveChanges();
                            dbContextTransaction.Commit();
                            result.RETURN_FLAG = true;
                            result.DESCRIPTION = $"減損報表資料刪除作業 ReportDate:{_reportDate}, FLOWID:{flowId} 成功";
                        }
                        else
                        {
                            result.DESCRIPTION = "刪除筆數為0";
                        }
                    }
                 ModC10Log(result, reportDate, Convert.ToInt16(ver)); //190620 PGE需求新增

                }
                catch (Exception ex)
                {
                    result.DESCRIPTION = ex.exceptionMessage();
                }
            }
            return result;
        }

        #endregion

        #region Excel 部分

        #region D54Excel
        public MSGReturnModel DownLoadExcel<T>(string type, string path, List<T> data)
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            result.DESCRIPTION = Message_Type.download_Fail
                .GetDescription(type, Message_Type.not_Find_Any.GetDescription());
            if (Excel_DownloadName.D54.ToString().Equals(type))
            {
                result.DESCRIPTION = FileRelated.DataTableToExcel(data.Cast<D54ViewModel>().ToList().ToDataTable(), path, Excel_DownloadName.D54);
                result.RETURN_FLAG = string.IsNullOrWhiteSpace(result.DESCRIPTION);
            }
            return result;
        }

        public MSGReturnModel DownLoadExcel2<T>(string type, string path, string path2, List<T> data)
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            result.DESCRIPTION = Message_Type.download_Fail
                .GetDescription(type, Message_Type.not_Find_Any.GetDescription());
            if (type == Excel_DownloadName.D63.ToString())
            {
                var D63 = data.Cast<D63ViewModel>().ToList();
                var D63Message = FileRelated.DataTableToExcel(D63.ToDataTable(), path, Excel_DownloadName.D63, new D63ViewModel().GetFormateTitles());
                var D64 = D63GetD64(D63);
                var D64Message = FileRelated.DataTableToExcel(D64.ToDataTable(), path2, Excel_DownloadName.D64, new D64ViewModel().GetFormateTitles());
                var _DESCRIPTION = string.Empty;
                if (!D63Message.IsNullOrWhiteSpace())
                    _DESCRIPTION += (D63Message + " ");
                if (!D64Message.IsNullOrWhiteSpace())
                    _DESCRIPTION += D64Message;
                result.DESCRIPTION = _DESCRIPTION;
                result.RETURN_FLAG = string.IsNullOrWhiteSpace(result.DESCRIPTION);
            }
            if (type == Excel_DownloadName.D65.ToString())
            {
                var D65 = data.Cast<D65ViewModel>().ToList();
                var D65Message = FileRelated.DataTableToExcel(D65.ToDataTable(), path, Excel_DownloadName.D65, new D65ViewModel().GetFormateTitles());
                var D66 = D65GetD66(D65);
                var D66Message = FileRelated.DataTableToExcel(D66.ToDataTable(), path2, Excel_DownloadName.D66, new D66ViewModel().GetFormateTitles());
                var _DESCRIPTION = string.Empty;
                if (!D65Message.IsNullOrWhiteSpace())
                    _DESCRIPTION += (D65Message + " ");
                if (!D66Message.IsNullOrWhiteSpace())
                    _DESCRIPTION += D66Message;
                result.DESCRIPTION = _DESCRIPTION;
                result.RETURN_FLAG = string.IsNullOrWhiteSpace(result.DESCRIPTION);
            }
            return result;
        }
        #endregion

        #endregion

        #region private function

        #region SaveModC10Log(190620 PGE需求新增)
        public void ModC10Log(MSGReturnModel result, string reportDt, int ver)
        {
            //DateTime reportdate = reportDt;
            if (result.RETURN_FLAG)
            {
                using (IFRS9DBEntities db = new IFRS9DBEntities())
                {
                    string sql = string.Empty;
                    sql += $@"UPDATE Transfer_CheckTable

                    Set TransferType = 'D',Error_Msg='減損報表資料刪除作業' 
                    WHERE ReportDate = '{reportDt.Replace("/", "-")}'
                    AND Version ={ver}
                    AND TransferType = 'Y'
                    AND File_Name ='{Table_Type.C10.ToString()}' ;";
                    db.Database.ExecuteSqlCommand(sql);
                }
            }
        }



        #endregion

        /// <summary>
        /// D6 複核系列程式
        /// </summary>
        /// <param name="Reference_Nbr"></param>
        /// <param name="Version"></param>
        /// <param name="Assessment_Result_Version"></param>
        /// <param name="type"></param>
        /// <param name="status"></param>
        /// <returns></returns>
        private MSGReturnModel D6ReviewMethod(
            string Reference_Nbr,
            int Assessment_Result_Version,
            Table_Type type,
            Evaluation_Status_Type status)
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            result.DESCRIPTION = Message_Type.already_Change.GetDescription();
            if (!(status == Evaluation_Status_Type.NotReview ||
                status  == Evaluation_Status_Type.Review ||
                status  == Evaluation_Status_Type.Reject ||
                status  == Evaluation_Status_Type.ReviewCompleted))
            {
                return result;
            }
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                if (type  == Table_Type.D63)
                {
                    var D63 = db.Bond_Quantitative_Resource.FirstOrDefault(x =>
                    x.Reference_Nbr == Reference_Nbr &&
                    x.Assessment_Result_Version == Assessment_Result_Version);
                    var D64 = db.Bond_Quantitative_Result.FirstOrDefault(x =>
                    x.Reference_Nbr == Reference_Nbr &&
                    x.Assessment_Result_Version == Assessment_Result_Version &&
                    x.Check_Item_Code.StartsWith(checkItemCode));
                    if (D63 != null && D64 != null)
                    {
                        D63.LastUpdate_User = _UserInfo._user;
                        D63.LastUpdate_Date = _UserInfo._date;
                        D63.LastUpdate_Time = _UserInfo._time;
                        switch (status)
                        {
                            case Evaluation_Status_Type.NotReview: //評估者取消該版本推送(取消複核)
                                if (D63.Audit_date != null) //複核者已經執行動作
                                    return result;
                                D63.Send_to_Auditor = null; //取消提交複核
                                D63.Send_Time = null;
                                break;
                            case Evaluation_Status_Type.Review: //評估者選擇該版本推送(等待複核)
                                D63.Send_to_Auditor = "Y"; //提交複核
                                D63.Send_Time = _UserInfo._date;
                                break;
                            case Evaluation_Status_Type.Reject: //複核者拒絕該版本(複核拒絕)
                                if (D63.Send_to_Auditor != "Y")
                                    return result;
                                D63.Auditor_Return = "Y";
                                D63.Audit_date = _UserInfo._date;
                                break;
                            case Evaluation_Status_Type.ReviewCompleted: //複核者通過該版本(複核通過)
                                if (D63.Send_to_Auditor != "Y")
                                    return result;
                                D63.Audit_date = _UserInfo._date;
                                var passConfirm = D64.Quantitative_Pass;
                                db.Bond_Quantitative_Resource.Where(x =>
                                x.Reference_Nbr == Reference_Nbr)
                                .ToList().ForEach(x =>
                                {
                                    x.Result_Version_Confirm = Assessment_Result_Version;
                                    x.Quantitative_Pass_Confirm = passConfirm;
                                    x.LastUpdate_User = _UserInfo._user;
                                    x.LastUpdate_Date = _UserInfo._date;
                                    x.LastUpdate_Time = _UserInfo._time;
                                });
                                break;
                        }
                    }
                    else
                    {
                        return result;
                    }
                }
                else if (type == Table_Type.D65)
                {
                    var D65 = db.Bond_Qualitative_Assessment.FirstOrDefault(x=>
                    x.Reference_Nbr == Reference_Nbr &&
                    x.Assessment_Result_Version == Assessment_Result_Version);
                    var D66 = db.Bond_Qualitative_Assessment_Result.FirstOrDefault(x =>
                    x.Reference_Nbr == Reference_Nbr &&
                    x.Assessment_Result_Version == Assessment_Result_Version &&
                    (x.Check_Item_Code.StartsWith(checkItemCode) || x.Check_Item_Code.StartsWith("Z")));
                    D65.LastUpdate_User = _UserInfo._user;
                    D65.LastUpdate_Date = _UserInfo._date;
                    D65.LastUpdate_Time = _UserInfo._time;
                    if (D65 != null && D66 != null)
                    {
                        switch (status)
                        {
                            case Evaluation_Status_Type.NotReview: //評估者取消該版本推送(取消複核)
                                if (D65.Audit_date != null) //複核者已經執行動作
                                    return result;
                                D65.Send_to_Auditor = null; //取消提交複核
                                D65.Send_Time = null;
                                break;
                            case Evaluation_Status_Type.Review: //評估者選擇該版本推送(等待複核)
                                D65.Send_to_Auditor = "Y"; //提交複核
                                D65.Send_Time = _UserInfo._date;
                                break;
                            case Evaluation_Status_Type.Reject: //複核者拒絕該版本(複核拒絕)
                                if (D65.Send_to_Auditor != "Y")
                                    return result;
                                D65.Auditor_Return = "Y";
                                D65.Audit_date = _UserInfo._date;
                                break;
                            case Evaluation_Status_Type.ReviewCompleted: //複核者通過該版本(複核通過)
                                if (D65.Send_to_Auditor != "Y")
                                    return result;
                                D65.Audit_date = _UserInfo._date;
                                var Assessment_Stage = D66.Assessment_Stage;
                                var passConfirm = Assessment_Stage == "2" ? D66.Qualitative_Pass_Stage2 :
                                                  Assessment_Stage == "3" ? D66.Qualitative_Pass_Stage3 :
                                                  "N";
                                db.Bond_Qualitative_Assessment.Where(x =>
                                x.Reference_Nbr == Reference_Nbr)
                                .ToList().ForEach(x =>
                                {
                                    x.Result_Version_Confirm = Assessment_Result_Version;
                                    x.Quantitative_Pass_Confirm = passConfirm;
                                    x.LastUpdate_User = _UserInfo._user;
                                    x.LastUpdate_Date = _UserInfo._date;
                                    x.LastUpdate_Time = _UserInfo._time;
                                });
                                break;
                        }

                    }
                    else
                    {
                        return result;
                    }
                }
                else
                {
                    return result;
                }
                try
                {
                    db.SaveChanges();
                    result.RETURN_FLAG = true;
                    var _DESCRIPTION = string.Empty;
                    switch (status)
                    {
                        case Evaluation_Status_Type.NotReview:
                            _DESCRIPTION = "該版本已取消提交!";
                            break;
                        case Evaluation_Status_Type.Review:
                            _DESCRIPTION = "該版本已提交!";
                            break;
                        case Evaluation_Status_Type.Reject:
                            _DESCRIPTION = "複核已退回!";
                            break;
                        case Evaluation_Status_Type.ReviewCompleted:
                            _DESCRIPTION = "複核已完成!";
                            break;
                    }
                    result.DESCRIPTION = _DESCRIPTION;
                }
                catch (DbUpdateException ex)
                {
                    result.DESCRIPTION = ex.exceptionMessage();
                }
            }
            return result;
        }

        #region Db 組成 D63ViewModel
        /// <summary>
        /// D63
        /// </summary>
        /// <param name="data"></param>
        /// <param name="flag">複核後選擇版本(已確認最後複核)</param>
        /// <param name="Users">Users</param>
        /// <returns></returns>
        private D63ViewModel DbToD63ViewModel(Bond_Quantitative_Resource data,List<IFRS9_User> Users)
        {
            return new D63ViewModel()
            {
                Status = data.Assessment_Sub_Kind == AssessmentSubKind_Type.Other.GetDescription() ?
                     Evaluation_Status_Type.ReviewCompleted.GetDescription() : string.Empty,
                Reference_Nbr = data.Reference_Nbr,
                Bond_Number = data.Bond_Number,
                Lots = data.Lots,
                Segment_Name = data.Segment_Name,
                Version = TypeTransfer.intNToString(data.Version),
                Report_Date = TypeTransfer.dateTimeNToString(data.Report_Date),
                Processing_Date = TypeTransfer.dateTimeNToString(data.Processing_Date),
                Assessment_Sub_Kind = data.Assessment_Sub_Kind,
                Net_Debt = TypeTransfer.doubleNToString(data.Net_Debt),
                Total_Asset = TypeTransfer.doubleNToString(data.Total_Asset),
                CFO = TypeTransfer.doubleNToString(data.CFO),
                Int_Expense = TypeTransfer.doubleNToString(data.Int_Expense),
                CET1_Capital_Ratio = TypeTransfer.doubleNToString(data.CET1_Capital_Ratio),
                Lerverage_Ratio = TypeTransfer.doubleNToString(data.Lerverage_Ratio),
                Liquiduty_Coverage_Ratio = TypeTransfer.doubleNToString(data.Liquiduty_Coverage_Ratio),
                Total_Equity = TypeTransfer.doubleNToString(data.Total_Equity),
                BS_TOT_ASSET = TypeTransfer.doubleNToString(data.BS_TOT_ASSET),
                RBC_Ratio = TypeTransfer.doubleNToString(data.RBC_Ratio),
                MCCSR_Ratio = TypeTransfer.doubleNToString(data.MCCSR_Ratio),
                Solvency_II_Ratio = TypeTransfer.doubleNToString(data.Solvency_II_Ratio),
                SHORT_AND_LONG_TERM_DEBT = TypeTransfer.doubleNToString(data.SHORT_AND_LONG_TERM_DEBT),
                Cash_and_Cash_Equivalents = TypeTransfer.doubleNToString(data.Cash_and_Cash_Equivalents),
                Short_Term_Debt = TypeTransfer.doubleNToString(data.Short_Term_Debt),
                ISSUER_AREA = data.ISSUER_AREA,
                IGS_INDEX = TypeTransfer.doubleNToString(data.IGS_INDEX),
                IGS_INDEX_Increase = TypeTransfer.doubleNToString(data.IGS_INDEX_Increase),
                External_Debt_Rate = TypeTransfer.doubleNToString(data.External_Debt_Rate),
                FC_Indexed_Debt_Rate = TypeTransfer.doubleNToString(data.FC_Indexed_Debt_Rate),
                CEIC_Value = TypeTransfer.doubleNToString(data.CEIC_Value),
                Short_term_Debt_Ratio = TypeTransfer.doubleNToString(data.Short_term_Debt_Ratio),
                GDP_Yearly_Avg = TypeTransfer.doubleNToString(data.GDP_Yearly_Avg),
                Quantitative_CLO_1 = TypeTransfer.doubleNToString(data.Quantitative_CLO_1),
                Quantitative_CLO_2 = TypeTransfer.doubleNToString(data.Quantitative_CLO_2),
                Quantitative_CLO_2_Threshold = TypeTransfer.doubleNToString(data.Quantitative_CLO_2_Threshold),
                Quantitative_CLO_3 = TypeTransfer.doubleNToString(data.Quantitative_CLO_3),
                Quantitative_CLO_3_Threshold = TypeTransfer.doubleNToString(data.Quantitative_CLO_3_Threshold),
                Quantitative_CLO_4 = TypeTransfer.doubleNToString(data.Quantitative_CLO_4),
                Quantitative_CLO_4_Threshold = TypeTransfer.doubleNToString(data.Quantitative_CLO_4_Threshold),
                Portfolio = data.Portfolio,
                Origination_Date = TypeTransfer.dateTimeNToString(data.Origination_Date),
                Portfolio_Name = data.Portfolio_Name,
                Assessment_Result_Version = TypeTransfer.intNToString(data.Assessment_Result_Version),
                Result_Version_Confirm_Flag =
                data.Assessment_Sub_Kind == AssessmentSubKind_Type.Other.GetDescription() ? "Y" : "N",
                Pass_Confirm_Flag = data.Assessment_Result_Version > 1 ? "Y" : "N",
                Quantitative_Pass_Confirm = data.Quantitative_Pass_Confirm,
                Assessor = data.Assessor,
                Assessor_Name = common.getUserName(Users,data.Assessor),
                Assessment_date = TypeTransfer.dateTimeNToString(data.Assessment_date),
                Auditor = data.Auditor,
                Auditor_Name = common.getUserName(Users, data.Auditor),
                Audit_date = TypeTransfer.dateTimeNToString(data.Audit_date),
                Send_to_Auditor = data.Send_to_Auditor,
                Send_Time = TypeTransfer.dateTimeNToString(data.Send_Time),
                Result_Version_Confirm = TypeTransfer.intNToString(data.Result_Version_Confirm),
                Auditor_Return = data.Auditor_Return           
            };
        }
        #endregion Db 組成 D63ViewModel

        #region Db 組成 D64ViewModel
        /// <summary>
        /// D64
        /// </summary>
        /// <param name="data"></param>
        /// <returns></returns>
        private D64ViewModel DbToD64ViewModel(Bond_Quantitative_Result data,string Assessor,List<Bond_Quantitative_Result_File> files)
        {
            return new D64ViewModel()
            {
                Reference_Nbr = data.Reference_Nbr,
                Report_Date = TypeTransfer.dateTimeNToString(data.Report_Date),
                Assessment_Kind = data.Assessment_Kind,
                Assessment_Sub_Kind = data.Assessment_Sub_Kind,
                Check_Item_Code = data.Check_Item_Code,
                Check_Item = data.Check_Item,
                Check_Item_Memo = data.Check_Item_Memo,
                Check_Reference = data.Check_Reference,
                Assessment_Result_Version = TypeTransfer.intNToString(data.Assessment_Result_Version),
                Check_Item_Value = TypeTransfer.doubleNToString(data.Check_Item_Value),
                Pass_Value = TypeTransfer.doubleNToString(data.Pass_Value),
                Quantitative_Pass = data.Quantitative_Pass,
                Check_Symbol = data.Check_Symbol,
                Threshold = TypeTransfer.doubleNToString(data.Threshold),
                Memo = data.Memo,
                Assessor = Assessor,
                Bond_Number = data.Bond_Number,
                Lots = data.Lots,
                Portfolio = data.Portfolio,
                Origination_Date = TypeTransfer.dateTimeNToString(data.Origination_Date),
                Portfolio_Name = data.Portfolio_Name,
                Assessment_Result = data.Assessment_Result,
                Assessment_Stage = data.Assessment_Stage,
                Version = data.Version.ToString(),
                FileCount = files.Count(x=>x.Check_Reference == data.Check_Reference && x.Status == "Y").ToString()
            };
        }
        #endregion Db 組成 D64ViewModel

        #region Db 組成 D65ViewModel
        /// <summary>
        /// D65
        /// </summary>
        /// <param name="data"></param>
        /// <param name="Users"></param>
        /// <returns></returns>
        private D65ViewModel DbToD65ViewModel(Bond_Qualitative_Assessment data, List<IFRS9_User> Users)
        {
            return new D65ViewModel()
            {
                Reference_Nbr = data.Reference_Nbr,
                Bond_Number = data.Bond_Number,
                Report_Date = data.Report_Date.ToString("yyyy/MM/dd"),
                Version = data.Version.ToString(),
                Assessment_Kind = data.Assessment_Kind,
                Assessment_Sub_Kind = data.Assessment_Sub_Kind,
                Assessment_Result_Version = data.Assessment_Result_Version.ToString(),
                Lots = data.Lots,
                Portfolio = data.Portfolio,
                Origination_Date = TypeTransfer.dateTimeNToString(data.Origination_Date),
                Portfolio_Name = data.Portfolio_Name,
                Auditor_Return = data.Auditor_Return,
                Send_to_Auditor = data.Send_to_Auditor,
                Send_Time = TypeTransfer.dateTimeNToString(data.Send_Time),
                Assessor = data.Assessor,
                Assessor_Name = common.getUserName(Users, data.Assessor),
                Assessment_date = TypeTransfer.dateTimeNToString(data.Assessment_date),
                Auditor = data.Auditor,
                Auditor_Name = common.getUserName(Users, data.Auditor),
                Audit_date = TypeTransfer.dateTimeNToString(data.Audit_date),
                Result_Version_Confirm = TypeTransfer.intNToString(data.Result_Version_Confirm),
                Quantitative_Pass_Confirm = data.Quantitative_Pass_Confirm,
                Pass_Confirm_Flag = data.Assessment_Result_Version > 1 ? "Y" : "N",
                Extra_Case = data.Extra_Case
            };
        }
        #endregion

        #region Db 組成 D66ViewModel
        /// <summary>
        /// D66
        /// </summary>
        /// <param name="data"></param>
        /// <returns></returns>
        private D66ViewModel DbToD66ViewModel(Bond_Qualitative_Assessment_Result data,string Assessor, List<Bond_Qualitative_Assessment_Result_File> files)
        {
            return new D66ViewModel()
            {
                Reference_Nbr = data.Reference_Nbr,
                Version = data.Version.ToString(),
                Report_Date = TypeTransfer.dateTimeToString(data.Report_Date),
                Assessment_Kind = data.Assessment_Kind,
                Assessment_Sub_Kind = data.Assessment_Sub_Kind,
                Assessment_Stage = data.Assessment_Stage,
                Check_Item_Code = data.Check_Item_Code,
                Check_Item = data.Check_Item,
                Check_Item_Memo = data.Check_Item_Memo,
                Check_Reference = data.Check_Reference,
                Assessment_Result_Version = data.Assessment_Result_Version.ToString(),
                Pass_Value = TypeTransfer.doubleNToString(data.Pass_Value),
                Qualitative_Pass_Stage2 = data.Qualitative_Pass_Stage2,
                Qualitative_Pass_Stage3 = data.Qualitative_Pass_Stage3,
                Check_Symbol = data.Check_Symbol,
                Threshold = TypeTransfer.doubleNToString(data.Threshold),
                Bond_Number = data.Bond_Number,
                Assessor = Assessor,
                FileCount = files.Count(x => x.Check_Reference == data.Check_Reference && x.Status == "Y").ToString()
            };
        }

        private string getResultVersionConfirmFlag(Evaluation_Status_Type status)
        {
            return ((status & Evaluation_Status_Type.ReviewCompleted) == Evaluation_Status_Type.ReviewCompleted) ? "Y" : "N"; 
        }

        #endregion

        /// <summary>
        /// D63 資料 轉換 D64 Check_Item_Value 量化指標值
        /// </summary>
        /// <param name="item_code">測試指標編號</param>
        /// <param name="model">D63</param>
        /// <returns></returns>
        private double? D64ItemValue(string item_code, Bond_Quantitative_Resource model)
        {
            //產業公司
            if (model.Assessment_Sub_Kind == AssessmentSubKind_Type.Industrial_Company.GetDescription())
            {
                //X11 => Net Debt/Total Asset
                if (item_code == "X11" && model.Net_Debt.HasValue && model.Total_Asset.HasValue)
                {
                    return model.Net_Debt.Value / model.Total_Asset.Value;
                }
                //X12 => Net Debt/CFO
                else if (item_code == "X12" && model.Net_Debt.HasValue && model.CFO.HasValue)
                {
                    return model.Net_Debt.Value / model.CFO.Value;
                }
                //X13 => (CFO/ Interest Expense)+1
                else if (item_code == "X13" && model.CFO.HasValue && model.Int_Expense.HasValue)
                {
                    return (model.CFO.Value / model.Int_Expense.Value) + 1;
                }
                //X14 => CFO
                else if (item_code == "X14" && model.CFO.HasValue)
                {
                    return model.CFO.Value;
                }
            }
            //金融債主順位債券 & 金融債次順位債券
            else if (model.Assessment_Sub_Kind == AssessmentSubKind_Type.Financial_Debts.GetDescription() ||
                model.Assessment_Sub_Kind == AssessmentSubKind_Type.Financial_Debt_Bond.GetDescription())
            {
                //金融債主順位債券
                if (model.Assessment_Sub_Kind == AssessmentSubKind_Type.Financial_Debts.GetDescription())
                {
                    //銀行核心一級資本適足率
                    if (item_code == "X21")
                    {
                        return model.CET1_Capital_Ratio;
                    }
                    //銀行財務槓桿比率
                    else if (item_code == "X22")
                    {
                        return model.Lerverage_Ratio;
                    }
                    //銀行流動性比率
                    else if (item_code == "X23")
                    {
                        return model.Liquiduty_Coverage_Ratio;
                    }
                }
                //金融債次順位債券
                else if (model.Assessment_Sub_Kind == AssessmentSubKind_Type.Financial_Debt_Bond.GetDescription())
                {
                    //銀行核心一級資本適足率
                    if (item_code == "X31")
                    {
                        return model.CET1_Capital_Ratio;
                    }
                    //銀行財務槓桿比率
                    else if (item_code == "X32")
                    {
                        return model.Lerverage_Ratio;
                    }
                    //銀行流動性比率
                    else if (item_code == "X33")
                    {
                        return model.Liquiduty_Coverage_Ratio;
                    }
                }
            }
            //壽險與產險公司
            else if (model.Assessment_Sub_Kind == AssessmentSubKind_Type.Life_Or_Property_Insurance_Company.GetDescription())
            {
                //資本水準_美國RBC_Ratio
                if (item_code == "X41")
                {
                    return model.RBC_Ratio;
                }
                //資本水準_加拿大MCCSR_Ratio
                else if (item_code == "X42")
                {
                    return model.MCCSR_Ratio;
                }
                //資本水準_歐洲Solvency_II_Ratio
                else if (item_code == "X43")
                {
                    return model.Solvency_II_Ratio;
                }
                //集團合併Total Equity / Total Asset
                else if (item_code == "X44" && model.Total_Equity.HasValue && model.BS_TOT_ASSET.HasValue)
                {
                    return model.Total_Equity.Value / model.BS_TOT_ASSET.Value;
                }
                //集團合併Financial Leverage (Total Debt / (Total Debt + Total Equity))
                //2018/06/27 加入 同時要有 Total_Equity & Total Debt 才可以計算
                else if (item_code == "X45" && model.SHORT_AND_LONG_TERM_DEBT.HasValue && model.Total_Equity.HasValue)
                {
                    //if (model.Total_Equity.HasValue)
                        return model.SHORT_AND_LONG_TERM_DEBT.Value / (model.SHORT_AND_LONG_TERM_DEBT.Value + model.Total_Equity.Value);
                    //else
                    //    return model.SHORT_AND_LONG_TERM_DEBT.Value / model.SHORT_AND_LONG_TERM_DEBT.Value;
                }
                //發行機構(流動性緩衝部位/一年內到期借款)
                else if (item_code == "X46" && model.Cash_and_Cash_Equivalents.HasValue && model.Short_Term_Debt.HasValue)
                {
                    return model.Cash_and_Cash_Equivalents.Value / model.Short_Term_Debt.Value;
                }
            }
            //主權政府債
            else if (model.Assessment_Sub_Kind == AssessmentSubKind_Type.Sovereign_Government_Debt.GetDescription())
            {
                //政府債務/GDP
                if (item_code == "X51" && model.IGS_INDEX.HasValue)
                {
                    return model.IGS_INDEX.Value;
                }
                //過去4年政府債務/GD增加數
                else if (item_code == "X52")
                {
                    return model.IGS_INDEX_Increase;
                }
                //外人持有政府債務/總政府債務
                else if (item_code == "X53")
                {
                    return model.External_Debt_Rate;
                }
                //外幣計價政府債務/總政府債務
                else if (item_code == "X54")
                {
                    return model.FC_Indexed_Debt_Rate;
                }
                //(經常帳+FDI淨流入)/GDP
                else if (item_code == "X55")
                {
                    return model.CEIC_Value;
                }
                //短期外債/外匯儲備
                else if (item_code == "X56")
                {
                    return model.Short_term_Debt_Ratio;
                }
                //過去四年平均經濟成長
                else if (item_code == "X57")
                {
                    return model.GDP_Yearly_Avg;
                }
            }
            //CLO
            else if (model.Assessment_Sub_Kind == AssessmentSubKind_Type.CLO.GetDescription())
            {
                //加權平均信用評級因子
                if (item_code == "X61")
                {
                    return model.Quantitative_CLO_1;
                }
                //違約資產金額*損失率.
                else if (item_code == "X62" && model.Quantitative_CLO_2.HasValue)
                {
                    return model.Quantitative_CLO_2.Value;
                }
                //資產池本金金額+現金
                else if (item_code == "X63" && model.Quantitative_CLO_3.HasValue)
                {
                    return model.Quantitative_CLO_3.Value;
                }
                //主順位有擔保貸款資產
                else if (item_code == "X64" && model.Quantitative_CLO_4.HasValue)
                {
                    return model.Quantitative_CLO_4.Value;
                }
            }
            return null;
        }

        /// <summary>
        /// D64 Pass_Value 通過給定值
        /// </summary>
        /// <param name="D61">D61</param>
        /// <param name="CheckItem">量化指標值</param>
        /// <param name="model">D63</param>
        /// <returns></returns>
        private double D64PassValue(Bond_Assessment_Parm D61, double CheckItem, Bond_Quantitative_Resource model)
        {
            bool flag = false;
            //產業公司 or 金融債主順位債券 or 金融債次順位債券 or 壽險與產險公司 or 主權政府債
            if (model.Assessment_Sub_Kind == AssessmentSubKind_Type.Industrial_Company.GetDescription() ||
                model.Assessment_Sub_Kind == AssessmentSubKind_Type.Financial_Debts.GetDescription() ||
                model.Assessment_Sub_Kind == AssessmentSubKind_Type.Financial_Debt_Bond.GetDescription() ||
                model.Assessment_Sub_Kind == AssessmentSubKind_Type.Life_Or_Property_Insurance_Company.GetDescription() ||
                model.Assessment_Sub_Kind == AssessmentSubKind_Type.Sovereign_Government_Debt.GetDescription() 
                )
            {
                //過去4年政府債務/GD增加數
                //政府債務/GDP < 120%。若政府債務/GDP>120%，但過去四年政府債務/GDP 比重，增加數未超過 20%，則可視為增加幅度尚屬溫和而通過測試。 
                if (D61.Check_Item_Code == "X52")
                {
                    flag = setPassValue(model.IGS_INDEX.Value, ">=", 1.2) &&
                           setPassValue(CheckItem, D61.Check_Symbol, D61.Threshold);
                }
                //2018/05/14 條件 X12 (Net Debt/CFO) 如果 Check_Item_Value 為負數 就是不通過
                else if (D61.Check_Item_Code == "X12" && (model.CFO.HasValue && model.CFO.Value < 0) ||
                         D61.Check_Item_Code == "X13" && (model.CFO.HasValue && model.CFO.Value < 0))
                {
                    flag = false;
                }
                else
                {
                    flag = setPassValue(CheckItem, D61.Check_Symbol, D61.Threshold);
                }
            }
            //CLO
            else if (model.Assessment_Sub_Kind == AssessmentSubKind_Type.CLO.GetDescription())
            {
                //加權平均信用評級因子
                if (D61.Check_Item_Code == "X61")
                {
                    flag = setPassValue(CheckItem, D61.Check_Symbol, D61.Threshold);
                }
                //違約資產金額*損失率.
                else if (D61.Check_Item_Code == "X62")
                {
                    flag = setPassValue(CheckItem, D61.Check_Symbol, model.Quantitative_CLO_2_Threshold);
                }
                //資產池本金金額+現金
                else if (D61.Check_Item_Code == "X63")
                {
                    flag = setPassValue(CheckItem, D61.Check_Symbol, model.Quantitative_CLO_3_Threshold);
                }
                //主順位有擔保貸款資產
                else if (D61.Check_Item_Code == "X64")
                {
                    flag = setPassValue(CheckItem, D61.Check_Symbol, model.Quantitative_CLO_4_Threshold);
                }
                else
                {
                    flag = setPassValue(CheckItem, D61.Check_Symbol, D61.Threshold);
                }
            }
            return flag ? D61.Pass_Value.HasValue ? D61.Pass_Value.Value : 0d : 0d;
        }

        /// <summary>
        /// D64 門檻
        /// </summary>
        /// <param name="D61">D61</param>
        /// <param name="model">D63</param>
        /// <returns></returns>
        private double? D64Threshold(Bond_Assessment_Parm D61, Bond_Quantitative_Resource model)
        {
            //評估次分類 為 CLO 的門檻 由使用者輸入
            if (model.Assessment_Sub_Kind == AssessmentSubKind_Type.CLO.GetDescription())
            {
                //較低順位券次之本金餘額 Quantitative_CLO_2_Threshold
                if (D61.Check_Item_Code == "X62")
                    return model.Quantitative_CLO_2_Threshold;
                //較低順位券次之本金餘額 Quantitative_CLO_3_Threshold
                if (D61.Check_Item_Code == "X63")
                    return model.Quantitative_CLO_3_Threshold;
                //資產池的80% Quantitative_CLO_4_Threshold
                if (D61.Check_Item_Code == "X64")
                    return model.Quantitative_CLO_4_Threshold;
            }
            //其他項目為設定在D61
            return D61.Threshold;
        }

        /// <summary>
        /// D64 是否通過量化評估
        /// </summary>
        /// <param name="D61">D61</param>
        /// <param name="PassValue">通過給定值</param>
        /// <returns></returns>
        private string D64QuantitativePass(Bond_Assessment_Parm D61, double Check_Item_Value,double PassValue,double? CFO)
        {
            //只有 Check_Item_Code like '%Pass_Count_' 才可以比對
            if (D61.Check_Item_Code.StartsWith(checkItemCode))
            {
                return setPassValue(Check_Item_Value, D61.Check_Symbol, D61.Threshold) ? "Y" : "N";
            }
            //2018/05/14 條件 X12 (Net Debt/CFO) & X13(CFO/ Interest Expense)+1 
            if (D61.Check_Item_Code.StartsWith("X12") && (CFO.HasValue && CFO.Value < 0) ||
                D61.Check_Item_Code.StartsWith("X13") && (CFO.HasValue && CFO.Value < 0))
                return "N";
            //2018/05/04 改為都比對
            return PassValue > 0 ? "Y" : "N";
        }

        /// <summary>
        /// D66 是否通過質化評估
        /// </summary>
        /// <param name="item_code">測試指標編號</param>
        /// <param name="PassValue">通過給定值</param>
        /// <param name="stage">評估階段</param>
        /// <param name="stageFlag">是否通過質化評估 2 or 3 判斷</param>
        /// <param name="Threshold">門檻</param>
        /// <param name="Check_Symbol">檢查條件</param>
        /// <returns></returns>
        private string D66QuantitativePass(string item_code, double PassValue, string stage,string stageFlag,double? Threshold,string Check_Symbol)
        {
            string N = "N";
            string Y = "Y";
            //Qualitative_Pass_Stage2 or Qualitative_Pass_Stage3 判斷 要等於 評估階段
            if (stage != stageFlag)
                return N;
            //評估階段 stage2
            if (stage == "2")
            {
                //只有 Check_Item_Code like '%Pass_Count_' 才可以比對
                if (item_code.StartsWith(checkItemCode))
                {
                    if (!Threshold.HasValue)
                        return N;
                    return setPassValue(PassValue, Check_Symbol, Threshold) ? Y : N;
                }
                //return N;
                //2018/05/04 改為都比對
                if (!Threshold.HasValue)
                    return N;
                //return setPassValue(PassValue, Check_Symbol, Threshold) ? Y : N;
                return PassValue > 0 ? "Y" : "N";
            }
            //評估階段 stage3 
            return PassValue == 1 ? Y : N;
        }

        /// <summary>
        /// 設定 D64 PassValue 通過給定值
        /// </summary>
        /// <param name="CheckItem">量化指標值</param>
        /// <param name="Check_Symbol">檢查條件</param>
        /// <param name="Threshold">門檻</param>
        /// <returns></returns>
        private bool setPassValue(double CheckItem, string Check_Symbol,double? Threshold)
        {
            var flag = false;
            if (!Threshold.HasValue || Check_Symbol.IsNullOrWhiteSpace())
                return flag;
            switch (Check_Symbol.Trim())
            {
                case "<":
                    flag = CheckItem < Threshold.Value;
                    break;
                case ">":
                    flag = CheckItem > Threshold.Value;
                    break;
                case "<=":
                    flag = CheckItem <= Threshold.Value;
                    break;
                case ">=":
                    flag = CheckItem >= Threshold.Value;
                    break;
                case "=":
                    flag = CheckItem == Threshold.Value;
                    break;
                case "!=":
                case "<>":
                    flag = CheckItem != Threshold.Value;
                    break;
            }
            return flag;
        }

        /// <summary>
        /// AssessmentSubKind to SelectOption
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        private SelectOption getAssessmentSubKindOption(string value)
        {
            return EnumUtil.GetValues<AssessmentSubKind_Type>()
                        .Where(x =>
                        value == x.GetDescription())
                        .Select(x => new SelectOption()
                        {
                            Text = x.GetDescription(),
                            Value = x.ToString()
                        }).FirstOrDefault();
        }

        //private string getImpairment_Stage(Bond_Qualitative_Assessment_Result data)
        //{
        //    if (data.Assessment_Result_Version != data.Result_Version_Confirm)
        //        return string.Empty;
        //    if (data.Assessment_Stage == "2" &&
        //        data.Assessment_Result_Version == data.Result_Version_Confirm)
        //    {
        //        if (data.Qualitative_Pass_Stage2 == "Y")
        //            return "1";
        //        return "2";
        //    }
        //    if (data.Assessment_Stage == "3" &&
        //        data.Assessment_Result_Version == data.Result_Version_Confirm)
        //    {
        //        if (data.Qualitative_Pass_Stage3 == "Y")
        //            return "1";
        //        return "3";
        //    }
        //    return string.Empty;
        //}

        private bool getD54Check(DateTime date)
        {
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                var _date = date.ToString("yyyy-MM-dd");
                return db.IFRS9_EL.AsNoTracking().Any(x => x.Report_Date == _date);
            }
        }

        /// <summary>
        /// D63 狀態
        /// </summary>
        /// <param name="D63Models"></param>
        /// <returns></returns>
        private Evaluation_Status_Type getD63Status(List<D63ViewModel> D63Models)
        {
            return
            getStatus(
                D63Models.Any(x => x.Result_Version_Confirm == x.Assessment_Result_Version), //複核完成
                D63Models.Any(x => x.Send_to_Auditor == "Y" && x.Auditor_Return.IsNullOrWhiteSpace()), //已提交複核(複核者未回覆)
                D63Models.Any(x => x.Auditor_Return == "Y"), //複核被退回請重新提交
                D63Models.Where(x => x.Auditor_Return != "Y").Count() >= 2 //尚未提交複核(已評估)
                );          
        }

        /// <summary>
        /// D65 狀態
        /// </summary>
        /// <param name="D65Models"></param>
        /// <returns></returns>
        private Evaluation_Status_Type getD65Status(List<D65ViewModel> D65Models)
        {
            return getStatus(
                D65Models.Any(x => x.Result_Version_Confirm == x.Assessment_Result_Version), //複核完成
                D65Models.Any(x => x.Send_to_Auditor == "Y" && x.Auditor_Return.IsNullOrWhiteSpace()), //已提交複核(複核者未回覆)
                D65Models.Any(x => x.Auditor_Return == "Y"), //複核被退回請重新提交
                D65Models.Where(x => x.Auditor_Return != "Y").Count() >= 2 //尚未提交複核(已評估)
                );
        }

        /// <summary>
        /// 獲取狀態中文
        /// </summary>
        /// <param name="ReviewCompleted">複核完成</param>
        /// <param name="Review">已提交複核(複核者未回覆)</param>
        /// <param name="NotReview">尚未複核(已評估)</param>
        /// <returns></returns>
        private Evaluation_Status_Type getStatus(bool ReviewCompleted,bool Review,bool Reject, bool NotReview)
        {
            Evaluation_Status_Type Status = Evaluation_Status_Type.NotAssess;
            if (ReviewCompleted)
                Status = Evaluation_Status_Type.ReviewCompleted;
            else if (Review)
                Status = Evaluation_Status_Type.Review;
            else if (NotReview)
                Status = Evaluation_Status_Type.NotReview;
            else if (Reject)
                Status = Evaluation_Status_Type.Reject;
            return Status;
        }

        /// <summary>
        /// D63 dl excel get D64 data
        /// </summary>
        /// <param name="datas"></param>
        /// <returns></returns>
        private List<D64ViewModel> D63GetD64(List<D63ViewModel> datas)
        {
            List<D64ViewModel> results = new List<D64ViewModel>();
            datas.ForEach(x =>
            {
                int ARV = Int32.Parse(x.Assessment_Result_Version);
                results.AddRange(getD64(x.Reference_Nbr,ARV));
            });
            return results;
        }

        /// <summary>
        /// D65 dl excel get D66 data
        /// </summary>
        /// <param name="datas"></param>
        /// <returns></returns>
        private List<D66ViewModel> D65GetD66(List<D65ViewModel> datas)
        {
            List<D66ViewModel> results = new List<D66ViewModel>();
            datas.ForEach(x =>
            {
                int ARV = Int32.Parse(x.Assessment_Result_Version);
                results.AddRange(getD66(x.Reference_Nbr, ARV));
            });
            return results;
        }

        private void AddD6CheckView(List<D6CheckViewModel> result, string _Job_Details, string Status, StringBuilder sb)
        {
            result.Add(new D6CheckViewModel()
            {
                Job_Details = _Job_Details,
                Status = Status,
                Details = sb.ToString()
            });
        }

        #endregion

        #region getD67
        public Tuple<bool, List<D67ViewModel>> getD67(D67ViewModel dataModel)
        {
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                if (db.Bond_Rating_Warning.Any())
                {
                    var data = from q in db.Bond_Rating_Warning.AsNoTracking()
                               select q;

                    if (dataModel.Report_Date.IsNullOrWhiteSpace() == false)
                    {
                        DateTime dateReportDate = DateTime.Parse(dataModel.Report_Date);
                        data = data.Where(x => x.Report_Date == dateReportDate);
                    }

                    //if (dataModel.Version.IsNullOrWhiteSpace() == false)
                    //{
                    //    int intVersion = int.Parse(dataModel.Version);
                    //    data = data.Where(x => x.Version == intVersion);
                    //}

                    data = data.Where(x => x.Bond_Number == dataModel.Bond_Number, dataModel.Bond_Number.IsNullOrWhiteSpace() == false)
                               .Where(x => x.Wraming_1_Ind == dataModel.Wraming_1_Ind, dataModel.Wraming_1_Ind.IsNullOrWhiteSpace() == false)
                               .Where(x => x.New_Ind == dataModel.New_Ind, dataModel.New_Ind.IsNullOrWhiteSpace() == false)
                               .Where(x => x.Rating_diff_Over_Ind == dataModel.Rating_diff_Over_Ind, dataModel.Rating_diff_Over_Ind.IsNullOrWhiteSpace() == false)
                               .Where(x => x.Wraming_2_Ind == dataModel.Wraming_2_Ind, dataModel.Wraming_2_Ind.IsNullOrWhiteSpace() == false);

                    //if (dataModel.IsComplete.IsNullOrWhiteSpace() == false)
                    //{
                    //    if (dataModel.IsComplete == "Y")
                    //    {
                    //        data = data.Where(x => x.Report_Date.ToString() != ""
                    //                            && x.Version.ToString() != ""
                    //                            && x.Bond_Number.ToString() != ""
                    //                            && x.Wraming_1_Ind.ToString() != ""
                    //                            && x.New_Ind.ToString() != ""
                    //                            && x.Rating_diff_Over_Ind.ToString() != ""
                    //                            && x.Wraming_2_Ind.ToString() != "");
                    //    }
                    //    else if (dataModel.IsComplete == "N")
                    //    {
                    //        data = data.Where(x => x.Report_Date.ToString() == ""
                    //                            || x.Version.ToString() == ""
                    //                            || x.Bond_Number.ToString() == ""
                    //                            || x.Wraming_1_Ind.ToString() == ""
                    //                            || x.New_Ind.ToString() == ""
                    //                            || x.Rating_diff_Over_Ind.ToString() == ""
                    //                            || x.Wraming_2_Ind.ToString() == "");
                    //    }
                    //}

                    return new Tuple<bool, List<D67ViewModel>>(data.Any(),
                                                               data.AsEnumerable()
                                                                   .Select(x =>
                                                                   {
                                                                       return DbToD67Model(x);
                                                                   }
                                                                  ).ToList()
                                                              );
                }
            }

            return new Tuple<bool, List<D67ViewModel>>(false, new List<D67ViewModel>());
        }
        #endregion

        #region Db 組成 D67ViewModel
        private D67ViewModel DbToD67Model(Bond_Rating_Warning data)
        {
            return new D67ViewModel()
            {
                Report_Date = data.Report_Date.ToString("yyyy/MM/dd"),
                Version = data.Version.ToString(),
                Bond_Number = data.Bond_Number,
                ISSUER = data.ISSUER,
                PRODUCT = data.PRODUCT,
                Product_Group_1 = data.Product_Group_1,
                Product_Group_2 = data.Product_Group_2,
                Amort_Amt_Tw = data.Amort_Amt_Tw.ToString(),
                GRADE_Warning_F = data.GRADE_Warning_F.ToString(),
                GRADE_Warning_D = data.GRADE_Warning_D.ToString(),
                Rating_diff_Over_F = data.Rating_diff_Over_F.ToString(),
                Rating_diff_Over_N = data.Rating_diff_Over_N.ToString(),
                Rating_Worse = data.Rating_Worse,
                Bond_Area = data.Bond_Area,
                PD_Grade = data.PD_Grade.ToString(),
                Wraming_1_Ind = data.Wraming_1_Ind,
                New_Ind = data.New_Ind,
                Observation_Month = data.Observation_Month.ToString(),
                Rating_diff_Over_Ind = data.Rating_diff_Over_Ind,
                Wraming_2_Ind = data.Wraming_2_Ind,
                Change_Memo = data.Change_Memo,
                Rating_Change_Memo = data.Rating_Change_Memo
            };
        }

        #endregion

        #region DownLoadD67Excel
        public MSGReturnModel DownLoadD67Excel(string path, List<D67ViewModel> dbDatas)
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            result.DESCRIPTION = Message_Type.download_Fail
                                 .GetDescription(Message_Type.not_Find_Any.GetDescription());

            if (dbDatas.Any())
            {
                DataTable datas = getD67ModelFromDb(dbDatas).Item1;

                result.DESCRIPTION = FileRelated.DataTableToExcel(datas, path, Excel_DownloadName.D67);
                result.RETURN_FLAG = string.IsNullOrWhiteSpace(result.DESCRIPTION);

                if (result.RETURN_FLAG)
                {
                    result.DESCRIPTION = Message_Type.download_Success.GetDescription();
                }
            }

            return result;
        }

        #endregion

        #region getD67ModelFromDb
        private Tuple<DataTable> getD67ModelFromDb(List<D67ViewModel> dbDatas)
        {
            DataTable dt = new DataTable();

            try
            {
                #region 組出資料

                dt.Columns.Add("評估基準日/報導日", typeof(object));
                dt.Columns.Add("資料版本", typeof(object));
                dt.Columns.Add("債券編號", typeof(object));
                dt.Columns.Add("發行人", typeof(object));
                dt.Columns.Add("債券產品別(揭露使用)", typeof(object));
                dt.Columns.Add("資產群組別(提供信用風險資產減損預警彙總表使用)", typeof(object));
                dt.Columns.Add("資產群組別(提供投資標的信用風險預警彙總表使用)", typeof(object));
                dt.Columns.Add("金融資產餘額(台幣)攤銷後之成本數(台幣_月底匯率)", typeof(object));
                dt.Columns.Add("預警標準_國外(BB-)", typeof(object));
                dt.Columns.Add("預警標準_國內為twBB-", typeof(object));
                dt.Columns.Add("降三級評等門檻＿國外（BB+）", typeof(object));
                dt.Columns.Add("降三級評等門檻＿國內（twBBB-）", typeof(object));
                dt.Columns.Add("債項最低評等內容", typeof(object));
                dt.Columns.Add("國內/國外", typeof(object));
                dt.Columns.Add("評等主標尺_原始", typeof(object));
                dt.Columns.Add("是否評等過低預警", typeof(object));
                dt.Columns.Add("是否新投資", typeof(object));
                dt.Columns.Add("連續追蹤月數", typeof(object));
                dt.Columns.Add("是否最近6個月內最近信評與原始信評的等級差異降評三級以上(含)", typeof(object));
                dt.Columns.Add("是否降三級以上且信評過低", typeof(object));
                dt.Columns.Add("本月變動", typeof(object));
                dt.Columns.Add("六個月內降評三級以上(含)之債券, 最早出現降評三級以上(含)之評估日最近信評(債項信評)", typeof(object));

                foreach (D67ViewModel item in dbDatas)
                {
                    var nrow = dt.NewRow();
                    nrow["評估基準日/報導日"] = item.Report_Date;
                    nrow["資料版本"] = item.Version;
                    nrow["債券編號"] = item.Bond_Number;
                    nrow["發行人"] = item.ISSUER;
                    nrow["債券產品別(揭露使用)"] = item.PRODUCT;
                    nrow["資產群組別(提供信用風險資產減損預警彙總表使用)"] = item.Product_Group_1;
                    nrow["資產群組別(提供投資標的信用風險預警彙總表使用)"] = item.Product_Group_2;
                    nrow["金融資產餘額(台幣)攤銷後之成本數(台幣_月底匯率)"] = item.Amort_Amt_Tw;
                    nrow["預警標準_國外(BB-)"] = item.GRADE_Warning_F;
                    nrow["預警標準_國內為twBB-"] = item.GRADE_Warning_D;
                    nrow["降三級評等門檻＿國外（BB+）"] = item.Rating_diff_Over_F;
                    nrow["降三級評等門檻＿國內（twBBB-）"] = item.Rating_diff_Over_N;
                    nrow["債項最低評等內容"] = item.Rating_Worse;
                    nrow["國內/國外"] = item.Bond_Area;
                    nrow["評等主標尺_原始"] = item.PD_Grade;
                    nrow["是否評等過低預警"] = item.Wraming_1_Ind;
                    nrow["是否新投資"] = item.New_Ind;
                    nrow["連續追蹤月數"] = item.Observation_Month;
                    nrow["是否最近6個月內最近信評與原始信評的等級差異降評三級以上(含)"] = item.Rating_diff_Over_Ind;
                    nrow["是否降三級以上且信評過低"] = item.Wraming_2_Ind;
                    nrow["本月變動"] = item.Change_Memo;
                    nrow["六個月內降評三級以上(含)之債券, 最早出現降評三級以上(含)之評估日最近信評(債項信評)"] = item.Rating_Change_Memo;
                    dt.Rows.Add(nrow);
                }

                #endregion 組出資料
            }
            catch
            {
            }

            return new Tuple<DataTable>(dt);
        }

        #endregion

        #region saveD67
        public MSGReturnModel saveD67(string reportDate)
        {
            MSGReturnModel result = new MSGReturnModel();

            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                try
                {
                    DateTime dateReportDate = DateTime.Parse(reportDate);
                    DateTime lastReportDate = dateReportDate.AddDays(-dateReportDate.Day);

                    var query = db.Bond_Account_Info.AsNoTracking()
                                  .Where(x => x.Report_Date == dateReportDate)
                                  .OrderByDescending(x => x.Version)
                                  .FirstOrDefault();

                    if (query == null)
                    {
                        result.RETURN_FLAG = false;
                        result.DESCRIPTION = "A41：Bond_Account_Info 沒有 " + reportDate + " 的資料";
                        return result;
                    }

                    int intVersion = TypeTransfer.intNToInt(query.Version);

                    var A41Data = db.Bond_Account_Info.AsNoTracking()
                                    .Where(x => x.Report_Date == dateReportDate
                                             && x.Version == intVersion)
                                    .GroupBy(x => new
                                                {
                                                    x.Report_Date,
                                                    x.Version,
                                                    x.Bond_Number,
                                                    x.ISSUER,
                                                    x.PRODUCT,
                                                    x.Bond_Aera
                                                }
                                            )
                                    .ToList();

                    var D72Data = db.SMF_Group.AsNoTracking().ToList();
                    if (D72Data.Any() == false)
                    {
                        result.RETURN_FLAG = false;
                        result.DESCRIPTION = "D72：SMF_Group 無資料";
                        return result;
                    }

                    var A41TheSameYearMonthData = db.Bond_Account_Info.AsNoTracking()
                                                                      .Where(x => x.Origination_Date.Value.Year == dateReportDate.Year
                                                                               && x.Origination_Date.Value.Month == dateReportDate.Month
                                                                               && x.Version == intVersion)
                                                                      .ToList();

                    var A57Data = db.Bond_Rating_Info.AsNoTracking().Where(x => x.Report_Date == dateReportDate
                                                                             && x.Version == intVersion
                                                                             //&& x.Rating_Type == "評估日最近信評"
                                                                             //&& x.Rating_Object == "債項"
                                                                          )
                                                                    .ToList();
                    //if (!A57Data.Any())
                    //{
                    //    result.RETURN_FLAG = false;
                    //    result.DESCRIPTION = reportDate + "、版本"+ intVersion.ToString() +"、評估日最近信評、債項，" + Environment.NewLine
                    //                         + "A57：Bond_Rating_Info 查無相關資料，無法做比對";
                    //    return result;
                    //}

                    if (!A57Data.Any())
                    {
                        result.RETURN_FLAG = false;
                        result.DESCRIPTION = reportDate + "、版本" + intVersion.ToString() + "，" + Environment.NewLine
                                             + "A57：Bond_Rating_Info 查無相關資料，無法做比對";
                        return result;
                    }

                    //var A51Data = db.Grade_Moody_Info.AsNoTracking()
                    //                .ToList();
                    //if (!A51Data.Any())
                    //{
                    //    result.RETURN_FLAG = false;
                    //    result.DESCRIPTION = "A51：Grade_Moody_Info 無資料，無法做比對";
                    //    return result;
                    //}

                    var A52Data = db.Grade_Mapping_Info.AsNoTracking()
                                    .Where(x => x.Rating_Org == "SP" || x.Rating_Org == "CW")
                                    .Where(x => x.IsActive == "Y")
                                    .ToList();
                    //if (!A52Data.Any())
                    //{
                    //    result.RETURN_FLAG = false;
                    //    result.DESCRIPTION = "A52：Grade_Mapping_Info 無 sp 的資料，無法做比對";
                    //    return result;
                    //}

                    DateTime beforeSixMonthDate = dateReportDate.AddMonths(-4).AddDays(-dateReportDate.AddMonths(-4).Day);
                    var D62BeforeSixMonthData = db.Bond_Basic_Assessment.AsNoTracking().Where(x => x.Report_Date >= beforeSixMonthDate
                                                                                                && x.Report_Date <= dateReportDate)
                                                                                       .ToList();

                    var A41BeforeReportDateData = db.Bond_Account_Info.AsNoTracking()
                                                    .Where(x => x.Report_Date == lastReportDate)
                                                    .OrderByDescending(x => x.Report_Date)
                                                    .ToList();

                    var D67BeforeReportDateData = db.Bond_Rating_Warning.AsNoTracking()
                                                    .Where(x => x.Report_Date <= lastReportDate)
                                                    .OrderByDescending(x => x.Report_Date)
                                                    .ToList();

                    var D67BeforeSixMonthData = db.Bond_Rating_Warning.AsNoTracking()
                                                  .Where(x => x.Report_Date >= beforeSixMonthDate
                                                           && x.Report_Date <= dateReportDate)
                                                  .ToList();

                    var ratingInfoSampleInfoList = db.Rating_Info_SampleInfo.AsNoTracking()
                                                  .Where(x => x.Report_Date == dateReportDate)
                                                  .ToList();

                    db.Database.ExecuteSqlCommand(string.Format(@"DELETE FROM Bond_Rating_Warning 
                                                                  WHERE Report_Date = '{0}'", reportDate));

                    string tempReport_Date = "";
                    string tempVersion = "";
                    string tempBond_Number = "";
                    string tempISSUER = "";
                    string tempPRODUCT = "";
                    string tempProduct_Group_1 = "";
                    string tempProduct_Group_2 = "";
                    string tempAmort_Amt_Tw = "";
                    string tempGRADE_Warning_F = "";
                    string tempGRADE_Warning_D = "";
                    string tempRating_diff_Over_F = "";
                    string tempRating_diff_Over_N = "";
                    string tempRating_Worse = "";
                    string tempBond_Area = "";
                    string tempPD_Grade = "";
                    string tempWraming_1_Ind = "";
                    string tempNew_Ind = "";
                    string tempObservation_Month = "";
                    string tempRating_diff_Over_Ind = "";
                    string tempWraming_2_Ind = "";
                    string tempChange_Memo = "";
                    string tempRating_Change_Memo = "";
                    StringBuilder sb = new StringBuilder();
                    for (int i = 0; i < A41Data.Count(); i++)
                    {
                        tempReport_Date = TypeTransfer.dateTimeNToString(A41Data[i].Key.Report_Date);
                        tempVersion = A41Data[i].Key.Version.ToString();
                        tempBond_Number = A41Data[i].Key.Bond_Number;
                        tempISSUER = A41Data[i].Key.ISSUER;
                        tempPRODUCT = A41Data[i].Key.PRODUCT;
                        tempProduct_Group_1 = "";
                        tempProduct_Group_2 = "";
                        tempAmort_Amt_Tw = "0";
                        tempGRADE_Warning_F = "13";
                        tempGRADE_Warning_D = "15";
                        tempRating_diff_Over_F = "11";
                        tempRating_diff_Over_N = "13";
                        tempRating_Worse = "";
                        tempBond_Area = A41Data[i].Key.Bond_Aera;
                        tempPD_Grade = "";
                        tempWraming_1_Ind = "";
                        tempNew_Ind = "";
                        tempObservation_Month = "";
                        tempRating_diff_Over_Ind = "";
                        tempWraming_2_Ind = "";
                        tempChange_Memo = "";
                        tempRating_Change_Memo = "";

                        var SMFGroup = D72Data.Where(x => x.PRODUCT == tempPRODUCT && x.Product_Group_kind == "Product_Group_1")
                                              .FirstOrDefault();
                        if (SMFGroup != null)
                        {
                            tempProduct_Group_1 = SMFGroup.Product_Group;
                        }

                        SMFGroup = D72Data.Where(x => x.PRODUCT == tempPRODUCT && x.Product_Group_kind == "Product_Group_2")
                                          .FirstOrDefault();
                        if (SMFGroup != null)
                        {
                            tempProduct_Group_2 = SMFGroup.Product_Group;
                        }

                        tempAmort_Amt_Tw = (from t in A41TheSameYearMonthData
                                            where t.Bond_Number == tempBond_Number
                                            select t.Amort_Amt_Tw).Sum().ToString();

                        var A57BondNumber = A57Data.Where(x => x.Bond_Number == tempBond_Number
                                               ).ToList();

                        tempPD_Grade = getD67PD_Grade(A57BondNumber.Where(x => x.Rating_Type == "評估日最近信評"
                                                                       && x.Rating_Object == "債項").ToList(), "");

                        if (tempPD_Grade.IsNullOrWhiteSpace() == false)
                        {
                            if (tempBond_Area.IsNullOrWhiteSpace() == false)
                            {
                                if (tempBond_Area == "國外")
                                {
                                    tempRating_Worse = getA52Rating(A52Data, "SP", tempPD_Grade);
                                }
                                else if (tempBond_Area == "國內")
                                {
                                    tempRating_Worse = getA52Rating(A52Data, "CW", tempPD_Grade);
                                }
                            }
                        }

                        if (tempRating_Worse.IsNullOrWhiteSpace() == true)
                        {
                            tempRating_Worse = getRating_WorseFromCOLLAT_TYP(ratingInfoSampleInfoList, tempBond_Number, reportDate);
                        }

                        if (tempRating_Worse == "SR UNSECURED")
                        {
                            tempPD_Grade = getD67PD_Grade(A57BondNumber.Where(x => x.Rating_Object == "發行人").ToList(), "Y");
                        }

                        if (tempPD_Grade.IsNullOrWhiteSpace() == false)
                        {
                            if (tempBond_Area.IsNullOrWhiteSpace() == false)
                            {
                                if (tempBond_Area == "國外")
                                {
                                    tempRating_Worse = getA52Rating(A52Data, "SP", tempPD_Grade);

                                    if (int.Parse(tempPD_Grade) >= int.Parse(tempGRADE_Warning_F))
                                    {
                                        tempWraming_1_Ind = "Y";
                                    }
                                }
                                else if (tempBond_Area == "國內")
                                {
                                    tempRating_Worse = getA52Rating(A52Data, "CW", tempPD_Grade);

                                    if (int.Parse(tempPD_Grade) >= int.Parse(tempGRADE_Warning_D))
                                    {
                                        tempWraming_1_Ind = "Y";
                                    }
                                }
                            }
                        }

                        if (A41BeforeReportDateData.Any() == false)
                        {
                            tempNew_Ind = "Y";
                        }
                        else
                        {
                            var oneA41BeforeReportDateData = A41BeforeReportDateData.FirstOrDefault();
                            if (oneA41BeforeReportDateData != null)
                            {
                                if (A41BeforeReportDateData.Where(x => x.Report_Date == oneA41BeforeReportDateData.Report_Date
                                                                    && x.Bond_Number == tempBond_Number).Any() == false)
                                {
                                    tempNew_Ind = "Y";
                                }
                            }
                        }

                        tempObservation_Month = D67BeforeReportDateData.Where(x => x.Bond_Number == tempBond_Number
                                                                                && x.Wraming_1_Ind == "Y")
                                                                       .Count().ToString();

                        var oneD62BeforeSixMonth = D62BeforeSixMonthData
                                                   .Where(x => x.Bond_Number == tempBond_Number && x.Curr_Ori_Rating_Diff > 2)
                                                   .OrderBy(x=>x.Report_Date)
                                                   .FirstOrDefault();
                        if (oneD62BeforeSixMonth != null)
                        {
                            tempRating_diff_Over_Ind = "Y";
                        }

                        if (tempPD_Grade.IsNullOrWhiteSpace() == false)
                        {
                            if (tempRating_diff_Over_Ind == "Y" 
                                && tempBond_Area == "國外"
                                && (int.Parse(tempPD_Grade) >= int.Parse(tempRating_diff_Over_F)))
                            {
                                tempWraming_2_Ind = "Y";
                            }

                            if (tempRating_diff_Over_Ind == "Y"
                                && tempBond_Area == "國內"
                                && (int.Parse(tempPD_Grade) >= int.Parse(tempRating_diff_Over_N)))
                            {
                                tempWraming_2_Ind = "Y";
                            }
                        }

                        if (tempNew_Ind == "Y")
                        {
                            tempChange_Memo = "新投資";
                        }
                        else
                        {
                            if (tempWraming_1_Ind == "Y")
                            {
                                var oneD67 = D67BeforeReportDateData.FirstOrDefault();
                                if (oneD67 != null)
                                {
                                    oneD67 = D67BeforeReportDateData.Where(x => x.Report_Date == oneD67.Report_Date
                                                                             && x.Bond_Number == tempBond_Number)
                                                                    .FirstOrDefault();
                                }

                                if (oneD67 != null)
                                {
                                    if (tempPD_Grade.IsNullOrWhiteSpace() == false && oneD67.PD_Grade.ToString().IsNullOrWhiteSpace() == false)
                                    {
                                        if (int.Parse(tempPD_Grade) > int.Parse(oneD67.PD_Grade.ToString()))
                                        {
                                            tempChange_Memo = tempRating_Worse + "下降";
                                        }
                                    }
                                }
                            }

                            if (A41TheSameYearMonthData.Any())
                            {
                                tempChange_Memo = "增加" + Math.Round(float.Parse(tempAmort_Amt_Tw) / 1000000, MidpointRounding.AwayFromZero).ToString();
                            }
                        }

                        if (tempWraming_2_Ind == "Y" && tempRating_diff_Over_Ind == "Y")
                        {
                            if (oneD62BeforeSixMonth != null)
                            {
                                var oneD67BeforeSixMonth = D67BeforeSixMonthData.Where(x => x.Bond_Number == tempBond_Number
                                                                                         && x.Report_Date == oneD62BeforeSixMonth.Report_Date)
                                                                                .FirstOrDefault();
                                if (oneD67BeforeSixMonth != null)
                                {
                                    tempRating_Change_Memo = oneD67BeforeSixMonth.Rating_Worse + "下降";
                                }
                            }
                        }

                        sb.Append(string.Format(@"
                                                   INSERT INTO Bond_Rating_Warning(Report_Date,
                                                                                   Version,
                                                                                   Bond_Number,
                                                                                   ISSUER,
                                                                                   PRODUCT,
                                                                                   Product_Group_1,
                                                                                   Product_Group_2,
                                                                                   Amort_Amt_Tw,
                                                                                   GRADE_Warning_F,
                                                                                   GRADE_Warning_D,
                                                                                   Rating_diff_Over_F,
                                                                                   Rating_diff_Over_N,
                                                                                   Rating_Worse,
                                                                                   Bond_Area,
                                                                                   PD_Grade,
                                                                                   Wraming_1_Ind,
                                                                                   New_Ind,
                                                                                   Observation_Month,
                                                                                   Rating_diff_Over_Ind,
                                                                                   Wraming_2_Ind,
                                                                                   Change_Memo,
                                                                                   Rating_Change_Memo)
                                                                         VALUES('{0}',{1},'{2}','{3}','{4}','{5}','{6}',{7},{8},{9},{10},
                                                                                 {11},'{12}','{13}',{14},'{15}','{16}',{17},'{18}','{19}','{20}',
                                                                                '{21}');",
                                                                                tempReport_Date,
                                                                                tempVersion,
                                                                                tempBond_Number,
                                                                                tempISSUER.Replace("'","''"),
                                                                                tempPRODUCT,
                                                                                tempProduct_Group_1,
                                                                                tempProduct_Group_2,
                                                                                tempAmort_Amt_Tw.IsNullOrWhiteSpace() == false ? tempAmort_Amt_Tw : "NULL",
                                                                                tempGRADE_Warning_F.IsNullOrWhiteSpace() == false ? tempGRADE_Warning_F : "NULL",
                                                                                tempGRADE_Warning_D.IsNullOrWhiteSpace() == false ? tempGRADE_Warning_D : "NULL",
                                                                                tempRating_diff_Over_F.IsNullOrWhiteSpace() == false ? tempRating_diff_Over_F : "NULL",
                                                                                tempRating_diff_Over_N.IsNullOrWhiteSpace() == false ? tempRating_diff_Over_N : "NULL",
                                                                                tempRating_Worse,
                                                                                tempBond_Area,
                                                                                tempPD_Grade.IsNullOrWhiteSpace() == false ? tempPD_Grade : "NULL",
                                                                                tempWraming_1_Ind,
                                                                                tempNew_Ind,
                                                                                tempObservation_Month.IsNullOrWhiteSpace() == false ? tempObservation_Month : "NULL",
                                                                                tempRating_diff_Over_Ind,
                                                                                tempWraming_2_Ind,
                                                                                tempChange_Memo,
                                                                                tempRating_Change_Memo));
                    }

                    db.Database.ExecuteSqlCommand(sb.ToString());

                    result.RETURN_FLAG = true;
                    result.DESCRIPTION = "完成";
                }
                catch (DbUpdateException ex)
                {
                    result.RETURN_FLAG = false;
                    result.DESCRIPTION = ex.Message;
                }
            }

            return result;
        }

        #endregion

        #region getD67PD_Grade
        private string getD67PD_Grade(List<Bond_Rating_Info> bondRatingInfoList, string isRTG_Bloomberg_Field)
        {
            string PD_Grade = "";

            if (isRTG_Bloomberg_Field == "Y")
            {
                bondRatingInfoList = bondRatingInfoList.Where(x => x.RTG_Bloomberg_Field == "RTG_MDY_SEN_UNSECURED_DEBT" 
                                                                || x.RTG_Bloomberg_Field == "RTG_FITCH_SEN_UNSECURED"
                                                         ).ToList();
            }

            var bondRatingInfo = bondRatingInfoList.OrderByDescending(x => x.PD_Grade)
                                                   .FirstOrDefault();
            if (bondRatingInfo != null)
            {
                PD_Grade = bondRatingInfo.PD_Grade.ToString();
            }

            return PD_Grade;
        }
        #endregion

        #region getA52Rating
        private string getA52Rating(List<Grade_Mapping_Info> gradeMappingInfoList, string Rating_Org, string PD_Grade)
        {
            string Rating = "";

            if (PD_Grade.IsNullOrWhiteSpace() == false)
            {
                var gradeMappingInfo = gradeMappingInfoList.Where(x => x.Rating_Org == Rating_Org
                                                                    && x.PD_Grade.ToString() == PD_Grade)
                                                           .ToList();
                if (gradeMappingInfo.Any() == true)
                {
                    int ratingSortNumber1 = getRatingSortNumber(Rating_Org, gradeMappingInfo[0].Rating);
                    Rating = gradeMappingInfo[0].Rating;
                    int ratingSortNumber2 = 0;
                    for (int j = 0; j < gradeMappingInfo.Count; j++)
                    {
                        ratingSortNumber2 = getRatingSortNumber(Rating_Org, gradeMappingInfo[j].Rating);

                        if (ratingSortNumber2 > ratingSortNumber1)
                        {
                            Rating = gradeMappingInfo[j].Rating;
                            ratingSortNumber1 = ratingSortNumber2;
                        }
                    }
                }
            }

            return Rating;
        }
        #endregion

        #region getRatingSortNumber
        private int getRatingSortNumber(string Rating_Org, string Rating)
        {
            int sortNumber = 0;

            if (Rating_Org == "CW")
            {
                switch (Rating)
                {
                    case "twAAA":
                        sortNumber = 1;
                        break;
                    case "twAA+":
                        sortNumber = 2;
                        break;
                    case "twAA":
                        sortNumber = 3;
                        break;
                    case "twAA-":
                        sortNumber = 4;
                        break;
                    case "twA+":
                        sortNumber = 5;
                        break;
                    case "twA":
                        sortNumber = 6;
                        break;
                    case "twA-":
                        sortNumber = 7;
                        break;
                    case "twBBB+":
                        sortNumber = 8;
                        break;
                    case "twBBB":
                        sortNumber = 9;
                        break;
                    case "twBBB-":
                        sortNumber = 10;
                        break;
                    case "twBB+":
                        sortNumber = 11;
                        break;
                    case "twBB":
                        sortNumber = 12;
                        break;
                    case "twBB-":
                        sortNumber = 13;
                        break;
                    case "twB+":
                        sortNumber = 14;
                        break;
                    case "twB":
                        sortNumber = 15;
                        break;
                    case "twB-":
                        sortNumber = 16;
                        break;
                    case "twCCC":
                        sortNumber = 17;
                        break;
                }
            }

            return sortNumber;
        }
        #endregion

        #region getRating_Worse
        private string getRating_WorseFromCOLLAT_TYP(List<Rating_Info_SampleInfo> ratingInfoSampleInfoList, string bondNumber, string reportDate)
        {
            string ratingWorse = "";

            DateTime dateReportDate = DateTime.Parse(reportDate);

            var ratingInfoSampleInfo = ratingInfoSampleInfoList.Where(x => x.Bond_Number == bondNumber
                                                                        && x.Report_Date == dateReportDate)
                                                               .FirstOrDefault();
            if (ratingInfoSampleInfo != null)
            {
                ratingWorse = ratingInfoSampleInfo.COLLAT_TYP;
            }

            return ratingWorse;
        }
        #endregion

    }
}