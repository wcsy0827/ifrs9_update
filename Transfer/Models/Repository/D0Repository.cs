using System;
using System.Collections.Generic;
using System.Data.Entity.Infrastructure;
using System.Linq;
using System.Text;
using Transfer.Controllers;
using Transfer.Models.Interface;
using Transfer.Utility;
using Transfer.ViewModels;
using static Transfer.Enum.Ref;

namespace Transfer.Models.Repository
{
    public class D0Repository : ID0Repository
    {
        #region 其他

        #endregion 其他

        #region Get GroupProduct By DebtType

        /// <summary>
        /// get Db Group_Product data by debtType
        /// </summary>
        /// <param name="debtType">1.房貸 4.債券</param>
        /// <returns></returns>
        public Tuple<bool, List<GroupProductViewModel>> getGroupProductByDebtType(string debtType)
        {
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                if (db.Group_Product.Any())
                {
                    var query = from q in db.Group_Product.AsNoTracking()
                                          .Where(x => x.Group_Product_Code.StartsWith(debtType))
                                select q;

                    return new Tuple<bool, List<GroupProductViewModel>>((query.Count() > 0 ? true : false),
                        query.AsEnumerable().OrderBy(x => x.Group_Product_Code)
                        .Select(x => { return DbToGroupProductViewModel(x); }).ToList());
                }
            }
            return new Tuple<bool, List<GroupProductViewModel>>(true, new List<GroupProductViewModel>());
        }

        #endregion Get GroupProduct By DebtType

        #region Db 組成 GroupProductViewModel

        /// <summary>
        /// Db 組成 GroupProductViewModel
        /// </summary>
        /// <param name="data"></param>
        /// <returns></returns>
        private GroupProductViewModel DbToGroupProductViewModel(Group_Product data)
        {
            return new GroupProductViewModel()
            {
                Group_Product_Code = data.Group_Product_Code,
                Group_Product_Name = data.Group_Product_Name
            };
        }

        #endregion Db 組成 GroupProductViewModel

        #region getD03All
        public Tuple<bool, List<D03ViewModel>> getD03All()
        {
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                if (db.Loan_Stage_Setting.Any())
                {
                    return new Tuple<bool, List<D03ViewModel>>
                                (
                                    true,
                                    (
                                        from q in db.Loan_Stage_Setting.AsNoTracking()
                                                    .AsEnumerable()
                                        select DbToD03Model(q)
                                    ).ToList()
                                );
                }
            }

            return new Tuple<bool, List<D03ViewModel>>(true, new List<D03ViewModel>());
        }
        #endregion

        #region DbToD03Model
        private D03ViewModel DbToD03Model(Loan_Stage_Setting data)
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

            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                var UserData = db.IFRS9_User.Where(x => x.User_Account == data.Rule_setter).FirstOrDefault();
                if (UserData != null)
                {
                    ruleSetterName = UserData.User_Name;
                }

                UserData = db.IFRS9_User.Where(x => x.User_Account == data.Auditor).FirstOrDefault();
                if (UserData != null)
                {
                    auditorName = UserData.User_Name;
                }
            }

            return new D03ViewModel()
            {
                Parm_ID = data.Parm_ID.ToString(),
                Stage1 = data.Stage1,
                Stage2 = data.Stage2,
                Stage3 = data.Stage3,
                Using_Start_Date = data.Using_Start_Date.ToString("yyyy/MM/dd"),
                Using_End_Date = data.Using_End_Date.ToString("yyyy/MM/dd"),
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

        #region  getD03
        public Tuple<bool, List<D03ViewModel>> getD03(D03ViewModel dataModel)
        {
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                if (db.Loan_Stage_Setting.Any())
                {
                    var query = db.Loan_Stage_Setting.AsNoTracking()
                                  .Where(x => x.Parm_ID.ToString() == dataModel.Parm_ID, dataModel.Parm_ID.IsNullOrWhiteSpace() == false)
                                  .Where(x => x.Status == dataModel.Status, dataModel.Status.IsNullOrWhiteSpace() == false)
                                  .Where(x => x.IsActive == dataModel.IsActive, dataModel.IsActive.IsNullOrWhiteSpace() == false);

                    return new Tuple<bool, List<D03ViewModel>>(query.Any(), query
                                                                            .AsEnumerable()
                                                                            .Select(x => { return DbToD03Model(x); }).ToList());
                }
            }

            return new Tuple<bool, List<D03ViewModel>>(false, new List<D03ViewModel>());
        }
        #endregion

        #region saveD03
        public MSGReturnModel saveD03(string actionType, D03ViewModel dataModel)
        {
            MSGReturnModel result = new MSGReturnModel();

            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                try
                {
                    if (actionType == "Add")
                    {
                        Loan_Stage_Setting addData = new Loan_Stage_Setting();

                        addData.Stage1 = dataModel.Stage1;
                        addData.Stage2 = dataModel.Stage2;
                        addData.Stage3 = dataModel.Stage3;
                        addData.Using_Start_Date = DateTime.Parse(dataModel.Using_Start_Date);
                        addData.Using_End_Date = DateTime.Parse(dataModel.Using_End_Date);
                        addData.Rule_setter = AccountController.CurrentUserInfo.Name;
                        addData.Rule_setting_Date = DateTime.Now.Date;
                        addData.Auditor = "";
                        addData.Audit_Date = null;
                        addData.Status = "0";
                        addData.IsActive = "";

                        db.Loan_Stage_Setting.Add(addData);
                    }
                    else if (actionType == "Modify")
                    {
                    }

                    db.SaveChanges(); //Save

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

        #region deleteD03
        public MSGReturnModel deleteD03(string parmID)
        {
            MSGReturnModel result = new MSGReturnModel();

            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                try
                {
                    Loan_Stage_Setting dataEdit = db.Loan_Stage_Setting
                                                    .Where(x => x.Parm_ID.ToString() == parmID)
                                                    .FirstOrDefault();

                    dataEdit.Rule_setter = AccountController.CurrentUserInfo.Name;
                    dataEdit.Rule_setting_Date = DateTime.Now.Date;
                    dataEdit.Auditor = "";
                    dataEdit.Audit_Date = null;
                    if (dataEdit.Status == "0" && dataEdit.IsActive.IsNullOrWhiteSpace() == true)
                    {
                        dataEdit.Status = "2";
                    }
                    else
                    {
                        dataEdit.Status = "0";
                    }
                    dataEdit.IsActive = "N";

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

        #region sendD03ToAudit
        public MSGReturnModel sendD03ToAudit(string parmID, string auditor)
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

                        Loan_Stage_Setting oldData = db.Loan_Stage_Setting
                                                       .Where(x => x.Parm_ID.ToString() == tempParmID)
                                                       .FirstOrDefault();
                        oldData.Auditor = auditor;
                        oldData.Status = "1";
                    }

                    db.SaveChanges();
                    result.RETURN_FLAG = true;
                }
                catch (Exception ex)
                {
                    result.RETURN_FLAG = false;
                    result.DESCRIPTION = ex.Message;
                }
            }

            return result;
        }
        #endregion

        #region D03Audit
        public MSGReturnModel D03Audit(string parmID, string status)
        {
            MSGReturnModel result = new MSGReturnModel();

            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                try
                {
                    string[] arrayParmID = parmID.Split(',');
                    string tempParmID = arrayParmID[0];
                    Loan_Stage_Setting oldData = db.Loan_Stage_Setting
                                                   .Where(x => x.Parm_ID.ToString() == tempParmID)
                                                   .FirstOrDefault();
                    for (int i = 0; i < arrayParmID.Length; i++)
                    {
                        tempParmID = arrayParmID[i];

                        oldData = db.Loan_Stage_Setting
                                    .Where(x => x.Parm_ID.ToString() == tempParmID)
                                    .FirstOrDefault();

                        oldData.Audit_Date = DateTime.Now.Date;
                        oldData.Status = status;

                        if (status == "2" && oldData.IsActive.IsNullOrWhiteSpace() == true)
                        {
                            oldData.IsActive = "Y";
                        }
                    }

                    db.SaveChanges();

                    result.RETURN_FLAG = true;
                }
                catch (Exception ex)
                {
                    result.RETURN_FLAG = false;
                    result.DESCRIPTION = ex.Message;
                }
            }

            return result;
        }
        #endregion

        #region Get Data

        /// <summary>
        /// get D05 all data
        /// </summary>
        /// <returns></returns>
        public Tuple<bool, List<D05ViewModel>> getD05All(string debtType)
        {
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                if (db.Group_Product_Code_Mapping.Any())
                {
                    return new Tuple<bool, List<D05ViewModel>>
                    (
                        true,
                        (
                            from q in db.Group_Product_Code_Mapping.AsNoTracking()
                            .Where(x => x.Group_Product_Code.StartsWith(debtType))
                            .AsEnumerable()
                            .OrderBy(x => x.Group_Product_Code)
                            .ThenBy(x => x.Product_Code)
                            select DbToD05ViewModel(q)
                        ).ToList()
                    );
                }
            }

            return new Tuple<bool, List<D05ViewModel>>(true, new List<D05ViewModel>());
        }

        /// <summary>
        /// get D05 data
        /// </summary>
        /// <returns></returns>
        public Tuple<bool, List<D05ViewModel>> getD05(string debtType, string groupProductCode, string productCode, string processingDate1, string processingDate2)
        {
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                if (db.Group_Product_Code_Mapping.Any())
                {
                    var query = from q in db.Group_Product_Code_Mapping.AsNoTracking()
                                .Where(x => x.Group_Product_Code.StartsWith(debtType))
                                select q;

                    if (!groupProductCode.IsNullOrWhiteSpace())
                    {
                        query = query.Where(x => x.Group_Product_Code.Contains(groupProductCode));
                    }

                    if (!productCode.IsNullOrWhiteSpace())
                    {
                        query = query.Where(x => x.Product_Code.Contains(productCode));
                    }

                    if (!processingDate1.IsNullOrWhiteSpace())
                    {
                        DateTime dProcessDate1 = DateTime.Parse(processingDate1);
                        query = query.Where(x => x.Processing_Date >= dProcessDate1);
                    }

                    if (!processingDate2.IsNullOrWhiteSpace())
                    {
                        DateTime dProcessDate2 = DateTime.Parse(processingDate2);
                        query = query.Where(x => x.Processing_Date <= dProcessDate2);
                    }

                    return new Tuple<bool, List<D05ViewModel>>((query.Count() > 0 ? true : false),
                        query.AsEnumerable().OrderBy(x => x.Group_Product_Code).ThenBy(x => x.Product_Code).Select(x => { return DbToD05ViewModel(x); }).ToList());
                }
            }

            return new Tuple<bool, List<D05ViewModel>>(false, new List<D05ViewModel>());
        }

        #endregion Get Data

        #region Db 組成 D05ViewModel

        /// <summary>
        /// Db 組成 D05ViewModel
        /// </summary>
        /// <param name="data"></param>
        /// <returns></returns>
        private D05ViewModel DbToD05ViewModel(Group_Product_Code_Mapping data)
        {
            return new D05ViewModel()
            {
                Group_Product_Code = data.Group_Product_Code,
                Group_Product = data.Group_Product.Group_Product_Name,
                Product_Code = data.Product_Code,
                Processing_Date = TypeTransfer.dateTimeNToString(data.Processing_Date)
            };
        }

        #endregion Db 組成 D05ViewModel

        #region Save D05

        /// <summary>
        /// D05 save db
        /// </summary>
        /// <param name="dataModel"></param>
        /// <returns></returns>
        public MSGReturnModel saveD05(string debtType, string actionType, D05ViewModel dataModel)
        {
            MSGReturnModel result = new MSGReturnModel();

            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                try
                {
                    if (debtType != dataModel.Group_Product_Code.Substring(0, 1))
                    {
                        result.RETURN_FLAG = false;
                        result.DESCRIPTION = "套用產品群代碼 第一個字必須是 " + debtType;
                        return result;
                    }

                    if (actionType == "Add")
                    {
                        if (db.Group_Product_Code_Mapping.Where(x => x.Group_Product_Code == dataModel.Group_Product_Code
                                                                  && x.Product_Code == dataModel.Product_Code).Count() > 0)
                        {
                            result.RETURN_FLAG = false;
                            result.DESCRIPTION = "資料重複： " + dataModel.Group_Product_Code + "、" + dataModel.Product_Code + " 已存在";
                            return result;
                        }

                        db.Group_Product_Code_Mapping.Add(
                        new Group_Product_Code_Mapping()
                        {
                            Group_Product_Code = dataModel.Group_Product_Code,
                            Product_Code = dataModel.Product_Code,
                            Processing_Date = DateTime.Parse(dataModel.Processing_Date)
                        });
                    }
                    else if (actionType == "Modify")
                    {
                        Group_Product_Code_Mapping oldData = db.Group_Product_Code_Mapping
                                                               .Where(x => x.Group_Product_Code == dataModel.Group_Product_Code
                                                                        && x.Product_Code == dataModel.Product_Code).FirstOrDefault();
                        oldData.Processing_Date = DateTime.Parse(dataModel.Processing_Date);
                    }

                    db.SaveChanges(); //Save
                    result.RETURN_FLAG = true;
                }
                catch (DbUpdateException ex)
                {
                    result.RETURN_FLAG = false;
                    result.DESCRIPTION = Message_Type
                                         .save_Fail.GetDescription("D05",
                                         $"message: {ex.Message}" +
                                         $", inner message {ex.InnerException?.InnerException?.Message}");
                }
            }

            return result;
        }

        #endregion Save D05

        #region Delete D05

        /// <summary>
        /// D05 delete db
        /// </summary>
        /// <param name="dataModel"></param>
        /// <returns></returns>
        public MSGReturnModel deleteD05(string groupProductCode, string productCode)
        {
            MSGReturnModel result = new MSGReturnModel();

            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                try
                {
                    var query = db.Group_Product_Code_Mapping.AsEnumerable().Where(x => x.Group_Product_Code == groupProductCode 
                                                                                     && x.Product_Code == productCode);
                    db.Group_Product_Code_Mapping.RemoveRange(query);

                    db.SaveChanges(); //Save
                    result.RETURN_FLAG = true;
                }
                catch (DbUpdateException ex)
                {
                    var sb = new StringBuilder();
                    sb.AppendLine($"{ex?.InnerException?.InnerException?.Message}");

                    //foreach (var eve in ex.Entries)
                    //{
                    //    sb.AppendLine($"Entity of type {eve.Entity.GetType().Name} in state {eve.State} could not be updated");
                    //}

                    result.RETURN_FLAG = false;
                    result.DESCRIPTION = "此項目有被其他地方使用到，不可刪除。" + sb.ToString();
                }
            }
            return result;
        }

        #endregion Delete D05

        #region Get D01

        /// <summary>
        /// get D01 all data
        /// </summary>
        /// <param name="debtType">1.房貸 4.債券</param>
        /// <returns></returns>
        public Tuple<bool, List<D01ViewModel>> getD01All(string debtType)
        {
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                if (db.Flow_Info.Any())
                {
                    return new Tuple<bool, List<D01ViewModel>>
                    (
                        true,
                        (
                            from q in db.Flow_Info.AsNoTracking()
                            .Where(x => x.Group_Product_Code.StartsWith(debtType)).AsEnumerable().OrderBy(x => x.PRJID).ThenBy(x => x.FLOWID)
                            select DbToD01ViewModel(q)
                        ).ToList()
                    );
                }
            }
            return new Tuple<bool, List<D01ViewModel>>(true, new List<D01ViewModel>());
        }

        /// <summary>
        /// get D01 data
        /// </summary>
        /// <param name="dataModel"></param>
        /// <returns></returns>
        public Tuple<bool, List<D01ViewModel>> getD01(D01ViewModel dataModel)
        {
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                if (db.Flow_Info.Any())
                {
                    var query = from q in db.Flow_Info.AsNoTracking()
                                .Where(x => x.Group_Product_Code.StartsWith(dataModel.DebtType))
                                select q;
                
                    if (dataModel.PRJID.IsNullOrEmpty() == false)
                    {
                        query = query.Where(x => x.PRJID.Contains(dataModel.PRJID));
                    }
                
                    if (dataModel.FLOWID.IsNullOrEmpty() == false)
                    {
                        query = query.Where(x => x.FLOWID.Contains(dataModel.FLOWID));
                    }
                
                    if (dataModel.Group_Product_Code.IsNullOrEmpty() == false)
                    {
                        query = query.Where(x => x.Group_Product_Code.Contains(dataModel.Group_Product_Code));
                    }
                
                    //DateTime tempDate;
                    //if (dataModel.Publish_Date.IsNullOrEmpty() == false)
                    //{
                    //    tempDate = DateTime.Parse(dataModel.Publish_Date);
                    //    query = query.Where(x => x.Publish_Date == tempDate);
                    //}

                    DateTime? publish_Start = null;
                    DateTime? publish_End = null;

                    DateTime _publish_Start = DateTime.MinValue;
                    DateTime _publish_End = DateTime.MinValue;
                    if (DateTime.TryParse(dataModel.Publish_Date, out _publish_Start))
                        publish_Start = _publish_Start;
                    if (DateTime.TryParse(dataModel.Publish_Date_End, out _publish_End))
                        publish_End = _publish_End;

                    query = query.Where(x => x.Publish_Date >= publish_Start, publish_Start != null)
                        .Where(x => x.Publish_Date <= publish_End, publish_End != null);

                    return new Tuple<bool, List<D01ViewModel>>((query.Count() > 0 ? true : false),
                        query.AsEnumerable().OrderBy(x => x.PRJID).ThenBy(x => x.FLOWID).Select(x => { return DbToD01ViewModel(x); }).ToList());
                }
            }
            return new Tuple<bool, List<D01ViewModel>>(false, new List<D01ViewModel>());
        }

        #endregion Get D01

        #region Db 組成 D01ViewModel

        /// <summary>
        /// Db 組成 D01ViewModel
        /// </summary>
        /// <param name="data"></param>
        /// <returns></returns>
        private D01ViewModel DbToD01ViewModel(Flow_Info data)
        {
            return new D01ViewModel()
            {
                PRJID = data.PRJID,
                FLOWID = data.FLOWID,
                Group_Product_Code = data.Group_Product_Code,
                Publish_Date = TypeTransfer.dateTimeNToString(data.Publish_Date),
                Apply_On_Date = TypeTransfer.dateTimeNToString(data.Apply_On_Date),
                Apply_Off_Date = TypeTransfer.dateTimeNToString(data.Apply_Off_Date),
                Issuer = data.Issuer,
                Issuer_Name = data.IFRS9_User.User_Name,
                Memo = data.Memo
            };
        }

        #endregion Db 組成 D01ViewModel

        #region Save D01

        /// <summary>
        /// D01 save db
        /// </summary>
        /// <param name="dataModel"></param>
        /// <returns></returns>
        public MSGReturnModel saveD01(D01ViewModel dataModel)
        {
            MSGReturnModel result = new MSGReturnModel();

            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                try
                {
                    if (dataModel.ActionType == "Add")
                    {
                        if (db.Flow_Info.Where(x => x.PRJID == dataModel.PRJID && x.FLOWID == dataModel.FLOWID).Count() > 0)
                        {
                            result.RETURN_FLAG = false;
                            result.DESCRIPTION = "資料重複：" + dataModel.PRJID + " " + dataModel.FLOWID + " 已存在";
                            return result;
                        }

                        Flow_Info addData = new Flow_Info();
                        addData.PRJID = dataModel.PRJID;
                        addData.FLOWID = dataModel.FLOWID;
                        addData.Group_Product_Code = dataModel.Group_Product_Code;
                        addData.Publish_Date = DateTime.Parse(dataModel.Publish_Date);
                        addData.Apply_On_Date = DateTime.Parse(dataModel.Apply_On_Date);
                        if (dataModel.Apply_Off_Date.IsNullOrWhiteSpace() == false)
                        {
                            addData.Apply_Off_Date = DateTime.Parse(dataModel.Apply_Off_Date);
                        }
                        addData.Issuer = dataModel.Issuer;
                        addData.Memo = dataModel.Memo;

                        db.Flow_Info.Add(addData);
                    }
                    else if (dataModel.ActionType == "Modify")
                    {
                        Flow_Info oldData = db.Flow_Info.Where(x => x.PRJID == dataModel.PRJID && x.FLOWID == dataModel.FLOWID).FirstOrDefault();
                        oldData.PRJID = dataModel.PRJID;
                        oldData.FLOWID = dataModel.FLOWID;
                        oldData.Group_Product_Code = dataModel.Group_Product_Code;
                        oldData.Publish_Date = DateTime.Parse(dataModel.Publish_Date);
                        oldData.Apply_On_Date = DateTime.Parse(dataModel.Apply_On_Date);
                        oldData.Apply_Off_Date = null;
                        if (dataModel.Apply_Off_Date.IsNullOrWhiteSpace() == false)
                        {
                            oldData.Apply_Off_Date = DateTime.Parse(dataModel.Apply_Off_Date);
                        }
                        oldData.Issuer = dataModel.Issuer;
                        oldData.Memo = dataModel.Memo;
                    }

                    db.SaveChanges(); //Save
                    result.RETURN_FLAG = true;
                }
                catch (DbUpdateException ex)
                {
                    result.RETURN_FLAG = false;
                    result.DESCRIPTION = Message_Type
                                         .save_Fail.GetDescription("D01",
                                         $"message: {ex.Message}" +
                                         $", inner message {ex.InnerException?.InnerException?.Message}");
                }
            }
            return result;
        }

        #endregion Save D01

        #region Delete D01

        /// <summary>
        /// D01 delete db
        /// </summary>
        /// <param name="prjid">專案名稱</param>
        /// <param name="flowid">流程名稱</param>
        /// <returns></returns>
        public MSGReturnModel deleteD01(string prjid, string flowid)
        {
            MSGReturnModel result = new MSGReturnModel();

            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                try
                {
                    var query = db.Flow_Info.AsEnumerable().Where(x => x.PRJID == prjid && x.FLOWID == flowid);
                    db.Flow_Info.RemoveRange(query);

                    db.SaveChanges(); //Save
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

        #endregion Delete D01

        #region getGroupProductAll

        /// <summary>
        /// get Group_Product all data
        /// </summary>
        /// <returns></returns>
        public Tuple<bool, List<GroupProductViewModel>> getGroupProductAll(string debtType)
        {
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                if (db.Group_Product.Any())
                {
                    return new Tuple<bool, List<GroupProductViewModel>>
                    (
                        true,
                        (
                            from q in db.Group_Product.AsNoTracking()
                            .Where(x => x.Group_Product_Code.StartsWith(debtType)).AsEnumerable()
                            select DbToGroupProductViewModel(q)
                        ).ToList()
                    );
                }
            }
            return new Tuple<bool, List<GroupProductViewModel>>(true, new List<GroupProductViewModel>());
        }

        #endregion getGroupProductAll

        #region getGroupProduct

        /// <summary>
        /// get GroupProduct data
        /// </summary>
        /// <returns></returns>
        public Tuple<bool, List<GroupProductViewModel>> getGroupProduct(string debtType, string groupProductCode, string groupProductName)
        {
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                if (db.Group_Product.Any())
                {
                    var query = from q in db.Group_Product.AsNoTracking()
                                .Where(x => x.Group_Product_Code.StartsWith(debtType))
                                select q;

                    if (groupProductCode != "")
                    {
                        query = query.Where(x => x.Group_Product_Code.Contains(groupProductCode));
                    }

                    if (groupProductName != "")
                    {
                        query = query.Where(x => x.Group_Product_Name.Contains(groupProductName));
                    }

                    return new Tuple<bool, List<GroupProductViewModel>>(query.Any(),
                        query.AsEnumerable().Select(x => { return DbToGroupProductViewModel(x); }).ToList());
                }
            }
            return new Tuple<bool, List<GroupProductViewModel>>(false, new List<GroupProductViewModel>());
        }

        #endregion getGroupProduct

        #region Save GroupProduct

        /// <summary>
        /// GroupProduct save db
        /// </summary>
        /// <param name="dataModel"></param>
        /// <returns></returns>
        public MSGReturnModel saveGroupProduct(string debtType, string actionType, GroupProductViewModel dataModel)
        {
            MSGReturnModel result = new MSGReturnModel();

            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                try
                {
                    if (debtType != dataModel.Group_Product_Code.Substring(0, 1))
                    {
                        result.RETURN_FLAG = false;
                        result.DESCRIPTION = "產品群代碼 第一個字必須是 " + debtType;
                        return result;
                    }

                    if (actionType == "Add")
                    {
                        if (db.Group_Product.Where(x => x.Group_Product_Code == dataModel.Group_Product_Code).Count() > 0)
                        {
                            result.RETURN_FLAG = false;
                            result.DESCRIPTION = "資料重複：" + dataModel.Group_Product_Code + " 已存在";
                            return result;
                        }

                        db.Group_Product.Add(
                        new Group_Product()
                        {
                            Group_Product_Code = dataModel.Group_Product_Code,
                            Group_Product_Name = dataModel.Group_Product_Name
                        });
                    }
                    else if (actionType == "Modify")
                    {
                        Group_Product oldData = db.Group_Product.Find(dataModel.Group_Product_Code);
                        oldData.Group_Product_Name = dataModel.Group_Product_Name;
                    }

                    db.SaveChanges(); //Save
                    result.RETURN_FLAG = true;
                }
                catch (DbUpdateException ex)
                {
                    result.RETURN_FLAG = false;
                    result.DESCRIPTION = Message_Type
                                         .save_Fail.GetDescription("GroupProduct",
                                         $"message: {ex.Message}" +
                                         $", inner message {ex.InnerException?.InnerException?.Message}");
                }
            }
            return result;
        }

        #endregion Save GroupProduct

        #region deleteGroupProduct
        public MSGReturnModel deleteGroupProduct(string groupProductCode)
        {
            MSGReturnModel result = new MSGReturnModel();
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                try
                {
                    if (db.Flow_Info.Where(x=>x.Group_Product_Code == groupProductCode).Any() == true)
                    {
                        result.RETURN_FLAG = false;
                        result.DESCRIPTION = "此項目已被 D01套用流程資訊 使用到，不可刪除。";
                        return result;
                    }

                    if (db.Group_Product_Code_Mapping.Where(x => x.Group_Product_Code == groupProductCode).Any() == true)
                    {
                        result.RETURN_FLAG = false;
                        result.DESCRIPTION = "此項目已被 D05套用產品組合代碼 使用到，不可刪除。";
                        return result;
                    }

                    var query = db.Group_Product.AsEnumerable().Where(x => x.Group_Product_Code == groupProductCode);
                    db.Group_Product.RemoveRange(query);

                    db.SaveChanges(); //Save
                    result.RETURN_FLAG = true;
                }
                catch (Exception ex)
                {
                    result.RETURN_FLAG = false;
                    result.DESCRIPTION = ex.Message;
                }
            }

            return result;
        }

        #endregion

        #region getD02
        public Tuple<bool, List<D02ViewModel>> getD02(D02ViewModel dataModel)
        {
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                if (db.IFRS9_Loan_Report.Any())
                {
                    string reportDate = DateTime.Parse(dataModel.Report_Date).ToString("yyyy-MM-dd");
                    string reportDate2 = DateTime.Parse(dataModel.Report_Date).ToString("yyyy/MM/dd");

                    var query = from q in db.IFRS9_Loan_Report.AsNoTracking()
                                            .Where(x => x.Report_Date == reportDate || x.Report_Date == reportDate2)
                                select q;

                    return new Tuple<bool, List<D02ViewModel>>(query.Any(),
                        query.AsEnumerable().OrderBy(x => x.Reference_Nbr).Select(x => { return DbToD02ViewModel(x); }).ToList());
                }
            }

            return new Tuple<bool, List<D02ViewModel>>(false, new List<D02ViewModel>());
        }
        #endregion

        #region DbToD02ViewModel
        private D02ViewModel DbToD02ViewModel(IFRS9_Loan_Report data)
        {
            return new D02ViewModel()
            {
                PRJID = data.PRJID,
                FLOWID = data.FLOWID,
                Report_Date = data.Report_Date,
                Processing_Date = data.Processing_Date,
                Product_Code = data.Product_Code,
                Reference_Nbr = data.Reference_Nbr,
                Lifetime_EL = data.Lifetime_EL.ToString(),
                Y1_EL = data.Y1_EL.ToString(),
                EL = data.EL.ToString(),
                Impairment_Stage = data.Impairment_Stage,
                Loan_Risk_Type = data.Loan_Risk_Type,
                Version = data.Version.ToString(),
                Principal = data.Principal.ToString(),
                Collection_Ind = data.Collection_Ind,
                Loan_Type = data.Loan_Type,
                Exposure = data.Exposure.ToString(),
                NO34RCV = data.NO34RCV,
                EIR = data.EIR.ToString(),
                Interest = data.Interest.ToString(),
                Current_LGD = data.Current_LGD.ToString(),
                PD = data.PD.ToString(),
                Interest_Receivable = data.Interest_Receivable.ToString(),
                Principal_EL = data.Principal_EL.ToString(),
                Interest_Receivable_EL = data.Interest_Receivable_EL.ToString()
            };
        }

        #endregion

        #region saveD02
        public MSGReturnModel saveD02(string reportDate)
        {
            MSGReturnModel result = new MSGReturnModel();

            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                try
                {
                    string reportDate1 = DateTime.Parse(reportDate).ToString("yyyy-MM-dd");
                    string reportDate2 = DateTime.Parse(reportDate).ToString("yyyy/MM/dd");
                    DateTime dateReportDate = DateTime.Parse(reportDate);

                    var A02 = db.Loan_Account_Info.AsNoTracking()
                                .Where(x => x.Report_Date == dateReportDate)
                                .ToList();
                    if (A02.Any() == false)
                    {
                        result.RETURN_FLAG = false;
                        result.DESCRIPTION = "A02：Loan_Account_Info 沒有 " + reportDate + " 的資料";
                        return result;
                    }

                    var C07 = db.EL_Data_Out.AsNoTracking()
                                            .Where(x => x.Report_Date == reportDate1 || x.Report_Date == reportDate2)
                                            .ToList();
                    if (C07.Any() == false)
                    {
                        result.RETURN_FLAG = false;
                        result.DESCRIPTION = "C07：EL_Data_Out 沒有 " + reportDate + " 的資料";
                        return result;
                    }

                    var A08 = db.Loan_Report_Info.AsNoTracking()
                                .Where(x => x.Report_Date == dateReportDate)
                                .ToList();
                    if (A08.Any() == false)
                    {
                        result.RETURN_FLAG = false;
                        result.DESCRIPTION = "A08：Loan_Report_Info 沒有 " + reportDate + " 的資料";
                        return result;
                    }

                    var A01 = db.Loan_IAS39_Info.AsNoTracking()
                                .Where(x => x.Report_Date == dateReportDate)
                                .ToList();
                    if (A01.Any() == false)
                    {
                        result.RETURN_FLAG = false;
                        result.DESCRIPTION = "A01：Loan_IAS39_Info 沒有 " + reportDate + " 的資料";
                        return result;
                    }

                    var C01 = db.EL_Data_In.AsNoTracking()
                                .Where(x => x.Report_Date == dateReportDate)
                                .ToList();
                    if (C01.Any() == false)
                    {
                        result.RETURN_FLAG = false;
                        result.DESCRIPTION = "C01：EL_Data_In 沒有 " + reportDate + " 的資料";
                        return result;
                    }

                    db.Database.ExecuteSqlCommand(string.Format(@"DELETE FROM IFRS9_Loan_Report 
                                                                  WHERE Report_Date = '{0}' OR Report_Date = '{1}'", reportDate1, reportDate2));

                    string sql = $@"
                                        INSERT INTO IFRS9_Loan_Report
                                                   ([PRJID]
                                                   ,[FLOWID]
                                                   ,[Report_Date]
                                                   ,[Processing_Date]
                                                   ,[Product_Code]
                                                   ,[Reference_Nbr]
                                                   ,[Lifetime_EL]
                                                   ,[Y1_EL]
                                                   ,[EL]
                                                   ,[Impairment_Stage]
                                                   ,[Loan_Risk_Type]
                                                   ,[Version]
                                                   ,[Principal]
                                                   ,[Collection_Ind]
                                                   ,[Loan_Type]
                                                   ,[Exposure]
                                                   ,[NO34RCV]
                                                   ,[EIR]
                                                   ,[Interest]
                                                   ,[Current_LGD]
                                                   ,[PD]
                                                   ,[Interest_Receivable]
                                                   ,[Principal_EL]
                                                   ,[Interest_Receivable_EL])
                                        SELECT A.PRJID
                                              ,A.FLOWID
	                                          ,A.Report_Date
	                                          ,A.Processing_Date
	                                          ,A.Product_Code
	                                          ,A.Reference_Nbr
	                                          ,A.Lifetime_EL
	                                          ,A.Y1_EL
	                                          ,A.EL
	                                          ,A.Impairment_Stage
	                                          ,B.Loan_Risk_Type
	                                          ,A.Version
	                                          ,C.Principal
	                                          ,D.Collection_Ind
	                                          ,(CASE WHEN A.Reference_Nbr LIKE '[A-Za-z]%' THEN '個人金融業務' ELSE '法人金融業務' END) Loan_Type
	                                          ,E.Exposure
	                                          ,C.NO34RCV
	                                          ,E.EIR
	                                          ,B.Interest
	                                          ,E.Current_LGD
	                                          ,A.PD
	                                          ,C.Interest_Receivable
	                                          ,(A.EL - (C.Interest_Receivable * A.PD * E.Current_LGD)) Principal_EL
	                                          ,(C.Interest_Receivable * A.PD * E.Current_LGD) Interest_Receivable_EL
                                        FROM
                                        (
	                                        SELECT * FROM EL_Data_Out
	                                        WHERE Report_Date = '{reportDate1}' OR Report_Date = '{reportDate2}'
                                        ) A
                                        JOIN
                                        (
	                                        SELECT * FROM Loan_Report_Info
	                                        WHERE Report_Date = '{reportDate}' 
                                        ) B ON A.Reference_Nbr = B.Reference_Nbr
                                        JOIN
                                        (
	                                        SELECT * FROM Loan_IAS39_Info
	                                        WHERE Report_Date = '{reportDate}' 
                                        ) C ON A.Reference_Nbr = C.Reference_Nbr
                                        JOIN
                                        (
	                                        SELECT * FROM Loan_Account_Info
	                                        WHERE Report_Date = '{reportDate}' 
                                        ) D ON A.Reference_Nbr = D.Reference_Nbr
                                        JOIN
                                        (
	                                        SELECT * FROM EL_Data_In
	                                        WHERE Report_Date = '{reportDate}' 
                                        ) E ON A.Reference_Nbr = E.Reference_Nbr                        
                                  ";

                    db.Database.ExecuteSqlCommand(sql);

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

        #region saveD02AfterKRisk
        public void saveD02AfterKRisk(string reportDate)
        {
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                try
                {
                    string reportDate1 = DateTime.Parse(reportDate).ToString("yyyy-MM-dd");
                    string reportDate2 = DateTime.Parse(reportDate).ToString("yyyy/MM/dd");
                    DateTime dateReportDate = DateTime.Parse(reportDate);

                    var A02 = db.Loan_Account_Info.AsNoTracking()
                                .Where(x => x.Report_Date == dateReportDate)
                                .ToList();
                    if (A02.Any() == false)
                    {
                        return;
                    }

                    var C07 = db.EL_Data_Out.AsNoTracking()
                                            .Where(x => x.Report_Date == reportDate1 || x.Report_Date == reportDate2)
                                            .ToList();
                    if (C07.Any() == false)
                    {
                        return;
                    }

                    var A08 = db.Loan_Report_Info.AsNoTracking()
                                .Where(x => x.Report_Date == dateReportDate)
                                .ToList();
                    if (A08.Any() == false)
                    {
                        return;
                    }

                    var A01 = db.Loan_IAS39_Info.AsNoTracking()
                                .Where(x => x.Report_Date == dateReportDate)
                                .ToList();
                    if (A01.Any() == false)
                    {
                        return;
                    }

                    var C01 = db.EL_Data_In.AsNoTracking()
                                .Where(x => x.Report_Date == dateReportDate)
                                .ToList();
                    if (C01.Any() == false)
                    {
                        return;
                    }

                    db.Database.ExecuteSqlCommand(string.Format(@"DELETE FROM IFRS9_Loan_Report 
                                                                  WHERE Report_Date = '{0}' OR Report_Date = '{1}'", reportDate1, reportDate2));

                    string sql = $@"
                                        INSERT INTO IFRS9_Loan_Report
                                                   ([PRJID]
                                                   ,[FLOWID]
                                                   ,[Report_Date]
                                                   ,[Processing_Date]
                                                   ,[Product_Code]
                                                   ,[Reference_Nbr]
                                                   ,[Lifetime_EL]
                                                   ,[Y1_EL]
                                                   ,[EL]
                                                   ,[Impairment_Stage]
                                                   ,[Loan_Risk_Type]
                                                   ,[Version]
                                                   ,[Principal]
                                                   ,[Collection_Ind]
                                                   ,[Loan_Type]
                                                   ,[Exposure]
                                                   ,[NO34RCV]
                                                   ,[EIR]
                                                   ,[Interest]
                                                   ,[Current_LGD]
                                                   ,[PD]
                                                   ,[Interest_Receivable]
                                                   ,[Principal_EL]
                                                   ,[Interest_Receivable_EL])
                                        SELECT A.PRJID
                                              ,A.FLOWID
	                                          ,A.Report_Date
	                                          ,A.Processing_Date
	                                          ,A.Product_Code
	                                          ,A.Reference_Nbr
	                                          ,A.Lifetime_EL
	                                          ,A.Y1_EL
	                                          ,A.EL
	                                          ,A.Impairment_Stage
	                                          ,B.Loan_Risk_Type
	                                          ,A.Version
	                                          ,C.Principal
	                                          ,D.Collection_Ind
	                                          ,(CASE WHEN A.Reference_Nbr LIKE '[A-Za-z]%' THEN '個人金融業務' ELSE '法人金融業務' END) Loan_Type
	                                          ,E.Exposure
	                                          ,C.NO34RCV
	                                          ,E.EIR
	                                          ,B.Interest
	                                          ,E.Current_LGD
	                                          ,A.PD
	                                          ,C.Interest_Receivable
	                                          ,(A.EL - (C.Interest_Receivable * A.PD * E.Current_LGD)) Principal_EL
	                                          ,(C.Interest_Receivable * A.PD * E.Current_LGD) Interest_Receivable_EL
                                        FROM
                                        (
	                                        SELECT * FROM EL_Data_Out
	                                        WHERE Report_Date = '{reportDate1}' OR Report_Date = '{reportDate2}'
                                        ) A
                                        JOIN
                                        (
	                                        SELECT * FROM Loan_Report_Info
	                                        WHERE Report_Date = '{reportDate}' 
                                        ) B ON A.Reference_Nbr = B.Reference_Nbr
                                        JOIN
                                        (
	                                        SELECT * FROM Loan_IAS39_Info
	                                        WHERE Report_Date = '{reportDate}' 
                                        ) C ON A.Reference_Nbr = C.Reference_Nbr
                                        JOIN
                                        (
	                                        SELECT * FROM Loan_Account_Info
	                                        WHERE Report_Date = '{reportDate}' 
                                        ) D ON A.Reference_Nbr = D.Reference_Nbr
                                        JOIN
                                        (
	                                        SELECT * FROM EL_Data_In
	                                        WHERE Report_Date = '{reportDate}' 
                                        ) E ON A.Reference_Nbr = E.Reference_Nbr                        
                                  ";

                    db.Database.ExecuteSqlCommand(sql);
                }
                catch (Exception ex)
                {
                }
            }
        }
        #endregion

    }
}