using Excel;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Entity.Infrastructure;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using Transfer.Models.Interface;
using Transfer.Utility;
using Transfer.ViewModels;
using static Transfer.Enum.Ref;

namespace Transfer.Models.Repository
{
    public class A7Repository : IA7Repository
    {
        #region 其他

        private List<string> A73Array = new List<string>() { "TM", "Default" };

        public A7Repository()
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

        #region Get Data

        #region Get Moody_Monthly_PD_Info(A71)

        /// <summary>
        /// Get A71 Data
        /// </summary>
        /// <returns></returns>
        public Tuple<bool, List<Moody_Tm_YYYY>> GetA71()
        {
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                if (db.Moody_Tm_YYYY.Any())
                {
                    return new Tuple<bool, List<Moody_Tm_YYYY>>
                        (true, db.Moody_Tm_YYYY.AsNoTracking().ToList());
                }
            }

            return new Tuple<bool, List<Moody_Tm_YYYY>>(false, new List<Moody_Tm_YYYY>());
        }

        #endregion Get Moody_Monthly_PD_Info(A71)

        #region Get Tm_Adjust_YYYY(A72)

        /// <summary>
        /// Get A72 Data
        /// </summary>
        /// <returns></returns>
        public Tuple<bool, List<object>> GetA72()
        {
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                if (db.Moody_Tm_YYYY.Any())
                {
                    List<object> odatas = new List<object>();
                    DataTable datas = getExhibit29ModelFromDb(db.Moody_Tm_YYYY.AsNoTracking().ToList()).Item1;
                    odatas.Add(datas.Columns.Cast<DataColumn>()
                         .Select(x => x.ColumnName)
                         .ToArray()); //第一列 由Columns 組成Title
                    for (var i = 0; i < datas.Rows.Count; i++)
                    {
                        List<string> str = new List<string>();
                        for (int j = 0; j < datas.Rows[i].ItemArray.Count(); j++)
                        {
                            if (datas.Columns[j].ToString().IndexOf("TM") > -1)
                            {
                                str.Add("\"" + datas.Columns[j] + "\":\"" + datas.Rows[i].ItemArray[j].ToString() + "\"");
                            }
                            else
                            {
                                str.Add("\"" + datas.Columns[j] + "\":" + datas.Rows[i].ItemArray[j].ToString());
                            }
                            //object 格式為 'column' : Rows.Data
                        }
                        odatas.Add(JsonConvert.DeserializeObject<IDictionary<string, object>>
                            ("{" + string.Join(",", str) + "}")); //第二列以後組成 object
                    }
                    if (odatas.Count > 2)
                    {
                        return new Tuple<bool, List<object>>(true, odatas);
                    }
                    else
                    {
                        return new Tuple<bool, List<object>>(false, new List<object>());
                    }
                }
                else
                {
                    return new Tuple<bool, List<object>>(false, new List<object>());
                }
            }
        }

        #endregion Get Tm_Adjust_YYYY(A72)

        #region Get GM_YYYY(A73)

        /// <summary>
        /// Get A73 Data
        /// </summary>
        /// <returns></returns>
        public Tuple<bool, List<object>> GetA73()
        {
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                if (db.Moody_Tm_YYYY.Any())
                {
                    List<object> odatas = new List<object>();
                    DataTable datas = getExhibit29ModelFromDb(db.Moody_Tm_YYYY.AsNoTracking().ToList()).Item1;
                    odatas.Add(datas.Columns.Cast<DataColumn>()
                         .Where(x => A73Array.Contains(x.ColumnName))
                         .Select(x => x.ColumnName)
                         .ToArray()); //第一列 由Columns 組成Title
                    for (var i = 0; i < datas.Rows.Count; i++)
                    {
                        List<string> str = new List<string>();
                        for (int j = 0; j < datas.Rows[i].ItemArray.Count(); j++)
                        {
                            if (A73Array.Contains(datas.Columns[j].ToString()))
                            {
                                if (datas.Columns[j].ToString().IndexOf("TM") > -1)
                                {
                                    str.Add("\"" + datas.Columns[j] + "\":\"" + datas.Rows[i].ItemArray[j].ToString() + "\"");
                                }
                                else
                                {
                                    str.Add("\"" + datas.Columns[j] + "\":" + datas.Rows[i].ItemArray[j].ToString());
                                }
                            }
                            //object 格式為 'column' : Rows.Data
                        }
                        odatas.Add(JsonConvert.DeserializeObject<IDictionary<string, object>>
                            ("{" + string.Join(",", str) + "}")); //第二列以後組成 object
                    }
                    if (odatas.Count > 2)
                    {
                        return new Tuple<bool, List<object>>(true, odatas);
                    }
                    else
                    {
                        return new Tuple<bool, List<object>>(false, new List<object>());
                    }
                }
                else
                {
                    return new Tuple<bool, List<object>>(false, new List<object>());
                }
            }

        }

        #endregion Get GM_YYYY(A73)

        #region Get Grade_Moody_Info(A51)

        /// <summary>
        /// Get A51 Data
        /// </summary>
        /// <returns></returns>
        public Tuple<bool, List<A51ViewModel>> GetA51(string year)
        {
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                var data = new A5Repository().getA51(Audit_Type.All,year).Item2;
                if (db.Grade_Moody_Info.Any(x=>x.Data_Year == year))
                {
                    return new Tuple<bool, List<A51ViewModel>>
                        (data.Any(), data);
                }
                else
                {
                    return new Tuple<bool, List<A51ViewModel>>(false, new List<A51ViewModel>());
                }
            }
        }

        #endregion Get Grade_Moody_Info(A51)

        #region get A51 Search Year

        /// <summary>
        /// get A51 year
        /// </summary>
        /// <returns></returns>
        public List<string> GetA51SearchYear()
        {
            List<string> result = new List<string>();
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                if (db.Grade_Moody_Info.Any())
                {
                    result.AddRange(db.Grade_Moody_Info.AsNoTracking()
                        .GroupBy(x => x.Data_Year)
                        .Select(x => x.FirstOrDefault())
                        .AsEnumerable()
                        .Select(x => (x.Data_Year + $"({new A5Repository().A51Status(x.Status)})")).OrderByDescending(x => x));
                }
                else
                {
                    result.Add(" ");
                }
            }
            return result;
        }

        #endregion get A51 Search Year

        #region 檢查可否上傳 A51
        public bool getA51SaveFlag(int year)
        {
            bool result = true;
            string _year = year.ToString();
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                if (db.Grade_Moody_Info.AsNoTracking().Any(x =>
                 x.Data_Year == _year &&
                 x.Status == "1"))
                {
                    result = false;
                }
            }
            return result;
        }
        #endregion


        #endregion Get Data

        #region save Db

        #region save Moody_Monthly_PD_Info(A71)

        /// <summary>
        /// Save  Moody_Monthly_PD_Info(A71)
        /// </summary>
        /// <param name="dataModel"></param>
        /// <returns></returns>
        public MSGReturnModel saveA71(List<Exhibit29Model> dataModel)
        {
            MSGReturnModel result = new MSGReturnModel();
            try
            {
                using (IFRS9DBEntities db = new IFRS9DBEntities())
                {
                    if (db.Moody_Tm_YYYY.Any())
                        db.Moody_Tm_YYYY.RemoveRange(db.Moody_Tm_YYYY.ToList()); //資料全刪除
                    int id = 1;
                    foreach (var item in dataModel)
                    {
                        db.Moody_Tm_YYYY.Add(
                            new Moody_Tm_YYYY()
                            {
                                Id = id,
                                From_To = item.From_To,
                                Aaa = TypeTransfer.stringToDoubleN(item.Aaa),
                                Aa1 = TypeTransfer.stringToDoubleN(item.Aa1),
                                Aa2 = TypeTransfer.stringToDoubleN(item.Aa2),
                                Aa3 = TypeTransfer.stringToDoubleN(item.Aa3),
                                A1 = TypeTransfer.stringToDoubleN(item.A1),
                                A2 = TypeTransfer.stringToDoubleN(item.A2),
                                A3 = TypeTransfer.stringToDoubleN(item.A3),
                                Baa1 = TypeTransfer.stringToDoubleN(item.Baa1),
                                Baa2 = TypeTransfer.stringToDoubleN(item.Baa2),
                                Baa3 = TypeTransfer.stringToDoubleN(item.Baa3),
                                Ba1 = TypeTransfer.stringToDoubleN(item.Ba1),
                                Ba2 = TypeTransfer.stringToDoubleN(item.Ba2),
                                Ba3 = TypeTransfer.stringToDoubleN(item.Ba3),
                                B1 = TypeTransfer.stringToDoubleN(item.B1),
                                B2 = TypeTransfer.stringToDoubleN(item.B2),
                                B3 = TypeTransfer.stringToDoubleN(item.B3),
                                Caa1 = TypeTransfer.stringToDoubleN(item.Caa1),
                                Caa2 = TypeTransfer.stringToDoubleN(item.Caa2),
                                Caa3 = TypeTransfer.stringToDoubleN(item.Caa3),
                                Ca_C = TypeTransfer.stringToDoubleN(item.Ca_C),
                                WR = TypeTransfer.stringToDoubleN(item.WR),
                                Default_Value = TypeTransfer.stringToDoubleN(item.Default),
                            });
                        id += 1;
                    }
                    db.SaveChanges(); //Save
                    result.RETURN_FLAG = true;
                    result.DESCRIPTION = Message_Type.save_Success.GetDescription(Table_Type.A71.ToString());
                }
            }
            catch (DbUpdateException ex)
            {
                result.RETURN_FLAG = false;
                result.DESCRIPTION = Message_Type
                                    .save_Fail.GetDescription(Table_Type.A71.ToString(),
                                    $"message: {ex.Message}" +
                                    $", inner message {ex.InnerException?.InnerException?.Message}");
            }
            return result;
        }

        #endregion save Moody_Monthly_PD_Info(A71)

        #region Save Tm_Adjust_YYYY(A72)

        /// <summary>
        /// Save  Tm_Adjust_YYYY(A72)
        /// </summary>
        /// <returns></returns>
        public MSGReturnModel saveA72()
        {
            MSGReturnModel result = new MSGReturnModel();
            try
            {
                using (IFRS9DBEntities db = new IFRS9DBEntities())
                {
                    if (db.Moody_Tm_YYYY.Any())
                    {
                        DataTable datas = getExhibit29ModelFromDb(db.Moody_Tm_YYYY.ToList()).Item1;
                        string cs = common.RemoveEntityFrameworkMetadata(string.Empty);
                        using (var conn = new SqlConnection(cs))
                        {
                            conn.Open();
                            var sqlStr1 = DropCreateTable(Table_Type.A72.GetDescription(), datas);
                            var sqlStr2 = CreateA7Table(Table_Type.A72.GetDescription(), datas);
                            using (var cmd = new SqlCommand(sqlStr1, conn))
                            {
                                Extension.NlogSet(sqlStr1);
                                int count = cmd.ExecuteNonQuery();
                            }
                            using (var cmd = new SqlCommand(sqlStr2, conn))
                            {
                                //conn.Open();
                                //SqlDataReader reader = cmd.ExecuteReader();
                                //while (reader.Read())
                                //{
                                //    flag = true;
                                //}
                                //reader.Close();
                                Extension.NlogSet(sqlStr2);
                                int count = cmd.ExecuteNonQuery();
                                if (datas.Rows.Count > 0 && datas.Rows.Count.Equals(count))
                                {
                                    result.RETURN_FLAG = true;
                                    result.DESCRIPTION = Message_Type.save_Success
                                        .GetDescription(Table_Type.A72.ToString());
                                }
                                else
                                {
                                    result.RETURN_FLAG = false;
                                    result.DESCRIPTION = Message_Type.save_Fail
                                        .GetDescription(Table_Type.A72.ToString(), "新增筆數有誤!");
                                }
                            }
                        }
                    }
                }
            }
            catch (SqlException ex)
            {
                result.RETURN_FLAG = false;
                result.DESCRIPTION = Message_Type
                        .save_Fail.GetDescription(Table_Type.A72.ToString(),
                        $"message: {ex.Message}" +
                        $", inner message {ex.InnerException?.InnerException?.Message}");
            }
            catch (Exception ex)
            {
                result.RETURN_FLAG = false;
                result.DESCRIPTION = Message_Type
                        .save_Fail.GetDescription(Table_Type.A72.ToString(),
                        $"message: {ex.Message}" +
                        $", inner message {ex.InnerException?.InnerException?.Message}");
            }
            return result;
        }

        #endregion Save Tm_Adjust_YYYY(A72)

        #region Save GM_YYYY(A73)

        /// <summary>
        /// Save  GM_YYYY(A73)
        /// </summary>
        /// <returns></returns>
        public MSGReturnModel saveA73()
        {
            MSGReturnModel result = new MSGReturnModel();
            try
            {
                using (IFRS9DBEntities db = new IFRS9DBEntities())
                {
                    if (db.Moody_Tm_YYYY.Any())
                    {
                        DataTable datas = getExhibit29ModelFromDb(db.Moody_Tm_YYYY.ToList()).Item1;
                        DataTable A73Datas = FromA72GetA73(datas);
                        string cs = common.RemoveEntityFrameworkMetadata(string.Empty);
                        using (var conn = new SqlConnection(cs))
                        {
                            conn.Open();
                            var sqlStr1 = DropCreateTable(Table_Type.A73.GetDescription(), A73Datas);
                            var sqlStr2 = CreateA7Table(Table_Type.A73.GetDescription(), A73Datas);
                            using (var cmd = new SqlCommand(sqlStr1, conn))
                            {
                                Extension.NlogSet(sqlStr1);
                                var count = cmd.ExecuteNonQuery();
                            }
                            using (var cmd = new SqlCommand(sqlStr2, conn))
                            {
                                Extension.NlogSet(sqlStr2);
                                int count = cmd.ExecuteNonQuery();
                                if (A73Datas.Rows.Count > 0 && A73Datas.Rows.Count.Equals(count))
                                {
                                    result.RETURN_FLAG = true;
                                    result.DESCRIPTION = Message_Type.save_Success
                                        .GetDescription(Table_Type.A73.ToString());
                                }
                                else
                                {
                                    result.RETURN_FLAG = false;
                                    result.DESCRIPTION = Message_Type.save_Fail
                                        .GetDescription(Table_Type.A73.ToString(), "新增筆數有誤!");
                                }
                            }
                        }
                    }
                }
            }
            catch (SqlException ex)
            {
                result.RETURN_FLAG = false;
                result.DESCRIPTION = Message_Type
                        .save_Fail.GetDescription(Table_Type.A73.ToString(),
                        $"message: {ex.Message}" +
                        $", inner message {ex.InnerException?.InnerException?.Message}");
            }
            catch (Exception ex)
            {
                result.RETURN_FLAG = false;
                result.DESCRIPTION = Message_Type
                        .save_Fail.GetDescription(Table_Type.A73.ToString(),
                        $"message: {ex.Message}" +
                        $", inner message {ex.InnerException?.InnerException?.Message}");
            }
            return result;
        }

        #endregion Save GM_YYYY(A73)

        #region Save Grade_Moody_Info(A51)

        /// <summary>
        /// Save  Grade_Moody_Info(A51)
        /// </summary>
        /// <param name="year"></param>
        /// <returns></returns>
        public MSGReturnModel saveA51(int Year)
        {
            MSGReturnModel result = new MSGReturnModel();
            try
            {
                using (IFRS9DBEntities db = new IFRS9DBEntities())
                {
                    if (db.Moody_Tm_YYYY.Any())
                    {
                        string year = Year.ToString();
                        if (db.Grade_Moody_Info.Any())
                            db.Grade_Moody_Info.RemoveRange(
                               db.Grade_Moody_Info.Where(x=>x.Data_Year == year)); //資料刪除當前年分

                        db.SaveChanges();

                        var A51Data = getExhibit29ModelFromDb(db.Moody_Tm_YYYY.ToList());
                        List<Grade_Moody_Info> A51s = (db.Moody_Tm_YYYY.AsEnumerable().
                            Select((x, y) => new Grade_Moody_Info
                            {
                                Rating = x.From_To,
                                Data_Year = year,
                            })).ToList();
                        A51s.Add(new Grade_Moody_Info() { Rating = "WR", Data_Year = year });
                        A51s.Add(new Grade_Moody_Info() { Rating = "Default", Data_Year = year });
                        int grade_Adjust = 1;
                        int PDGrade = 1;
                        List<string> alreadyNum = new List<string>();
                        foreach (Grade_Moody_Info item in A51s)
                        {
                            string rating_Adjust = string.Empty;
                            foreach (var col in A51Data.Item2)
                            {
                                if (col.Value.Contains(item.Rating))
                                {
                                    rating_Adjust = col.Key + "_" + col.Value.Last();
                                }
                            }
                            if (!string.IsNullOrWhiteSpace(rating_Adjust)) //合併欄位情況
                            {
                                if (alreadyNum.Contains(rating_Adjust)) //與上一筆一樣是合併欄位
                                {
                                    grade_Adjust -= 1; //Grade_Adjust 不變
                                }
                                else
                                {
                                    alreadyNum.Add(rating_Adjust); //新的合併欄位
                                }
                            }
                            item.Grade_Adjust = grade_Adjust;
                            item.Moodys_PD = string.IsNullOrWhiteSpace(rating_Adjust) ?
                                A51Data.Item1.AsEnumerable().Where(x => x.Field<string>("TM") == item.Rating)
                                .Select(x => Convert.ToDouble(x.Field<string>("Default"))).FirstOrDefault() :
                                A51Data.Item1.AsEnumerable().Where(x => x.Field<string>("TM") == rating_Adjust)
                                .Select(x => Convert.ToDouble(x.Field<string>("Default"))).FirstOrDefault();
                            item.PD_Grade = PDGrade;
                            item.Rating_Adjust = rating_Adjust.Replace("_", "~");
                            grade_Adjust += 1;
                            PDGrade += 1;
                        }
                        A51s.ForEach(x =>
                        {
                            x.Create_User = _UserInfo._user;
                            x.Create_Date = _UserInfo._date;
                            x.Create_Time = _UserInfo._time;
                        });
                        db.Grade_Moody_Info.AddRange(A51s);
                        db.SaveChanges(); //Save
                        result.RETURN_FLAG = true;
                        result.DESCRIPTION = Message_Type.save_Success
                                              .GetDescription(Table_Type.A51.ToString());
                    }
                }
            }
            catch (DbUpdateException ex)
            {
                result.RETURN_FLAG = false;
                result.DESCRIPTION = Message_Type
                        .save_Fail.GetDescription(Table_Type.A51.ToString(),
                        $"message: {ex.Message}" +
                        $", inner message {ex.InnerException?.InnerException?.Message}");
            }
            return result;
        }

        #endregion Save Grade_Moody_Info(A51)

        /// <summary>
        /// 選擇A51複核者
        /// </summary>
        /// <param name="year">年份</param>
        /// <param name="Auditor">複核者</param>
        /// <returns></returns>
        public MSGReturnModel AssessmentA51(string year, string Auditor)
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                if (!common.getAssessmentInfo(Assessment_Type.B.GetDescription(),
                    Table_Type.A51.ToString(),
                    SetAssessmentType.Presented).Any(x => x.User_Account == _UserInfo._user))
                {
                    result.DESCRIPTION = Message_Type.none_Send_Audit_Authority.GetDescription();
                    return result;
                }
                if (!common.getAssessmentInfo(Assessment_Type.B.GetDescription(),
                    Table_Type.A51.ToString(),
                    SetAssessmentType.Auditor).Any(x => x.User_Account == Auditor))
                {
                    result.DESCRIPTION = Message_Type.none_Audit_Authority.GetDescription();
                    return result;
                }
                var A51s = db.Grade_Moody_Info.Where(x => x.Data_Year == year).ToList();
                if (A51s.Any())
                {
                    A51s.ForEach(x =>
                    {
                        x.Auditor = Auditor;
                        x.LastUpdate_User = _UserInfo._user;
                        x.LastUpdate_Date = _UserInfo._date;
                        x.LastUpdate_Time = _UserInfo._time;
                    });

                    try
                    {
                        db.SaveChanges();
                        result.RETURN_FLAG = true;
                        result.DESCRIPTION = Message_Type.send_To_Audit_Success.GetDescription();
                    }
                    catch (Exception ex)
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

        /// <summary>
        /// A51複核者只能變更 啟用 or 暫不啟用
        /// </summary>
        /// <param name="type">啟用 or 暫不啟用</param>
        /// <param name="year">年份</param>
        /// <param name="Auditor_Reply">複核訊息</param>
        /// <returns></returns>
        public MSGReturnModel AuditA51(Audit_Type type,string year,string Auditor_Reply)
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                var A51s = db.Grade_Moody_Info.Where(x => x.Data_Year == year).ToList();
                if (A51s.Any())
                {
                    var first = A51s.First();
                    if (first.Auditor != _UserInfo._user)
                    {
                        result.DESCRIPTION = $"複核者錯誤應該為{_UserInfo._user}";
                        return result;
                    }
                    if ((first.Status == "1" && type == Audit_Type.Enable) ||
                       (first.Status == "2" && type == Audit_Type.TempDisabled)
                       )
                    {
                        result.DESCRIPTION = $"年度:{first.Data_Year}的資料已經是{type.GetDescription()}狀態,故無修改.";
                    }
                    else
                    {
                        if (type == Audit_Type.Enable)
                        {
                            var _oldA51s = db.Grade_Moody_Info.Where(x => x.Status == "1");
                            foreach (var _oldA51 in _oldA51s)
                            {
                                _oldA51.Status = "3"; //停用
                                _oldA51.Auditor = _UserInfo._user;
                                _oldA51.Audit_date = _UserInfo._date;
                                _oldA51.LastUpdate_User = _UserInfo._user;
                                _oldA51.LastUpdate_Date = _UserInfo._date;
                                _oldA51.LastUpdate_Time = _UserInfo._time;
                            }

                            var D68s = db.Risk_Parm.Where(x => x.IsActive == "Y").ToList();
                            var _System = "System";
                            D68s.ForEach(x =>
                            {
                                Risk_Parm D68Add = x.ModelConvert<Risk_Parm, Risk_Parm>();
                                D68Add.Data_Year = first.Data_Year;
                                D68Add.Rule_setter = _System;
                                D68Add.Rule_setting_Date = _UserInfo._date;
                                D68Add.Grade_Adjust = TypeTransfer.intNToInt(resetGradeAdjust(A51s, D68Add.PD_Grade));
                                db.Risk_Parm.Add(D68Add);
                                x.IsActive = "N";
                            });
                            var D70s = db.Watching_List_Parm.Where(x => x.IsActive == "Y").ToList();
                            D70s.ForEach(x => {
                                Watching_List_Parm D70Add = x.ModelConvert<Watching_List_Parm, Watching_List_Parm>();
                                D70Add.Data_Year = first.Data_Year;
                                D70Add.Rule_setter = _System;
                                D70Add.Rule_setting_Date = _UserInfo._date;
                                D70Add.Rating_Threshold_Map_Grade_Adjust = resetGradeAdjust(A51s, D70Add.Rating_Threshold_Map_Grade);
                                D70Add.Rating_from_Map_Grade_Adjust = resetGradeAdjust(A51s, D70Add.Rating_from_Map_Grade);
                                D70Add.Rating_To_Map_Grade_Adjust = resetGradeAdjust(A51s, D70Add.Rating_To_Map_Grade);
                                db.Watching_List_Parm.Add(D70Add);
                                x.IsActive = "N";
                            });
                            var D71s = db.Warning_List_Parm.Where(x => x.IsActive == "Y").ToList();
                            D71s.ForEach(x =>
                            {
                                Warning_List_Parm D71Add = x.ModelConvert<Warning_List_Parm, Warning_List_Parm>();
                                D71Add.Data_Year = first.Data_Year;
                                D71Add.Rule_setter = _System;
                                D71Add.Rule_setting_Date = _UserInfo._date;
                                D71Add.Rating_Threshold_Map_Grade_Adjust = resetGradeAdjust(A51s, D71Add.Rating_Threshold_Map_Grade);
                                D71Add.Rating_from_Map_Grade_Adjust = resetGradeAdjust(A51s, D71Add.Rating_from_Map_Grade);
                                D71Add.Rating_To_Map_Grade_Adjust = resetGradeAdjust(A51s, D71Add.Rating_To_Map_Grade);
                                db.Warning_List_Parm.Add(D71Add);
                                x.IsActive = "N";
                            });
                        }
                        foreach (var A51 in A51s)
                        {
                            A51.Status = getA51Status(type);
                            A51.Auditor = _UserInfo._user;
                            A51.Auditor_Reply = Auditor_Reply;
                            A51.Audit_date = _UserInfo._date;
                            A51.LastUpdate_User = _UserInfo._user;
                            A51.LastUpdate_Date = _UserInfo._date;
                            A51.LastUpdate_Time = _UserInfo._time;
                        }
                        try
                        {
                            db.SaveChanges();
                            result.RETURN_FLAG = true;
                            result.DESCRIPTION = Message_Type.Audit_Success.GetDescription();
                        }
                        catch (Exception ex)
                        {
                            result.DESCRIPTION = ex.exceptionMessage();
                        }
                    }
                }
                else
                {
                    result.DESCRIPTION = Message_Type.not_Find_Any.GetDescription();
                }
            }
            return result;
        }

        #endregion save Db

        #region Excel 部分

        #region Excel 資料轉成 Exhibit29Model

        /// <summary>
        /// Excel 資料轉成 Exhibit29Model
        /// </summary>
        /// <param name="pathType">Excel 副檔名</param>
        /// <param name="stream"></param>
        /// <returns></returns>
        public Tuple<int, List<Exhibit29Model>> getExcel(string pathType, Stream stream)
        {
            DataSet resultData = new DataSet();
            List<Exhibit29Model> dataModel = new List<Exhibit29Model>();
            int year = 0;
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

                if (resultData.Tables[0].Rows.Count > 2) //判斷有無資料
                {
                    var _datas = resultData.Tables[0].AsEnumerable();
                    var title = TypeTransfer.objToString(_datas.AsDataView().Table.Columns[0]);
                    year = TypeTransfer.stringToInt(title.Substring(title.LastIndexOf("-") + 1).Trim());
                    if (year == 0)
                        year = DateTime.Now.Year - 1;
                    dataModel = (from q in _datas.Skip(1)
                                 select getExhibit29Model(q)).ToList();
                    //skip(1) 為排除Excel Title列那行(參數可調)
                    dataModel = dataModel.Take(dataModel.Count - 1).ToList();
                    //排除最後一筆 為 * Data in percent 的註解
                }
            }
            catch
            { }
            return new Tuple<int, List<Exhibit29Model>>(year, dataModel);
        }

        #endregion Excel 資料轉成 Exhibit29Model

        #region 下載 Excel

        /// <summary>
        /// 下載 Excel
        /// </summary>
        /// <param name="type">(A72.A73)</param>
        /// <param name="path">下載位置</param>
        /// <returns></returns>
        public MSGReturnModel DownLoadExcel(string type, string path)
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            result.DESCRIPTION = Message_Type.download_Fail
                .GetDescription(type, Message_Type.not_Find_Any.GetDescription());
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                if (db.Moody_Tm_YYYY.Any())
                {
                    DataTable datas = getExhibit29ModelFromDb(db.Moody_Tm_YYYY.AsNoTracking().ToList()).Item1;
                    if (Excel_DownloadName.A72.ToString().Equals(type))
                    {
                        result.DESCRIPTION = FileRelated.DataTableToExcel(datas, path, Excel_DownloadName.A72);
                        result.RETURN_FLAG = string.IsNullOrWhiteSpace(result.DESCRIPTION);
                    }
                    if (Excel_DownloadName.A73.ToString().Equals(type))
                    {
                        DataTable newData = FromA72GetA73(datas); //要組新的 Table
                        if (newData != null) //有資料
                        {
                            result.DESCRIPTION = FileRelated.DataTableToExcel(newData, path, Excel_DownloadName.A73);
                            result.RETURN_FLAG = string.IsNullOrWhiteSpace(result.DESCRIPTION);
                        }
                        else
                        {
                            result.DESCRIPTION = Message_Type.download_Fail.GetDescription(type, Message_Type.not_Find_Any.GetDescription());
                        }
                    }
                    if (result.RETURN_FLAG)
                        result.DESCRIPTION = Message_Type.download_Success.GetDescription(type);
                }
            }
            return result;
        }

        #endregion 下載 Excel

        #endregion Excel 部分

        #region Private Function

        private int? resetGradeAdjust(List<Grade_Moody_Info> A51s,int? PD_Grade)
        {
            if (PD_Grade.HasValue)
                return A51s.FirstOrDefault(x => x.PD_Grade == PD_Grade)?.Grade_Adjust;
            return null;
        }

        /// <summary>
        /// 複核者只能變更 啟用 or 暫不啟用
        /// </summary>
        /// <param name="type"></param>
        /// <returns></returns>
        private string getA51Status(Audit_Type type) 
        {
            if (type == Audit_Type.Enable)
                return "1";
            return "2";
        }

        #region datarow 組成 Exhibit29Model

        /// <summary>
        /// datarow 組成 Exhibit29Model
        /// </summary>
        /// <param name="item">DataRow</param>
        /// <returns>Exhibit29Model</returns>
        private Exhibit29Model getExhibit29Model(DataRow item)
        {
            return new Exhibit29Model()
            {
                From_To = TypeTransfer.objToString(item[0]),
                Aaa = TypeTransfer.objToString(item[1]),
                Aa1 = TypeTransfer.objToString(item[2]),
                Aa2 = TypeTransfer.objToString(item[3]),
                Aa3 = TypeTransfer.objToString(item[4]),
                A1 = TypeTransfer.objToString(item[5]),
                A2 = TypeTransfer.objToString(item[6]),
                A3 = TypeTransfer.objToString(item[7]),
                Baa1 = TypeTransfer.objToString(item[8]),
                Baa2 = TypeTransfer.objToString(item[9]),
                Baa3 = TypeTransfer.objToString(item[10]),
                Ba1 = TypeTransfer.objToString(item[11]),
                Ba2 = TypeTransfer.objToString(item[12]),
                Ba3 = TypeTransfer.objToString(item[13]),
                B1 = TypeTransfer.objToString(item[14]),
                B2 = TypeTransfer.objToString(item[15]),
                B3 = TypeTransfer.objToString(item[16]),
                Caa1 = TypeTransfer.objToString(item[17]),
                Caa2 = TypeTransfer.objToString(item[18]),
                Caa3 = TypeTransfer.objToString(item[19]),
                Ca_C = TypeTransfer.objToString(item[20]),
                WR = TypeTransfer.objToString(item[21]),
                Default = TypeTransfer.objToString(item[22]),
            };
        }

        #endregion datarow 組成 Exhibit29Model

        #region DB Moody_Tm_YYYY 組成 DataTable

        /// <summary>
        ///  DB Moody_Tm_YYYY 組成 DataTable
        /// </summary>
        /// <param name="dbDatas"></param>
        /// <returns></returns>
        private Tuple<DataTable, Dictionary<string, List<string>>> getExhibit29ModelFromDb(List<Moody_Tm_YYYY> dbDatas)
        {
            DataTable dt = new DataTable();
            //超過的筆對紀錄 => string 開始的參數,List<string>要相加的參數
            Dictionary<string, List<string>> overData =
                new Dictionary<string, List<string>>();
            try
            {
                #region 找出錯誤的參數

                string errorKey = string.Empty; //錯誤起始欄位
                string last_FromTo = string.Empty; //上一個 From_To
                double last_value = 0d; //上一個default的參數(#)
                foreach (Moody_Tm_YYYY item in dbDatas) //第一次迴圈先抓出不符合的項目
                {
                    double now_default_value =  //目前的default 參數
                        TypeTransfer.doubleNToDouble(item.Default_Value);
                    if (now_default_value >= last_value) //下一筆比上一筆大(正常情況)
                    {
                        if (!errorKey.IsNullOrWhiteSpace()) //假如上一筆是超過的參數
                        {
                            errorKey = string.Empty; //把錯誤Flag 取消掉(到上一筆為止)
                        }

                        #region 把現在的參數寄到最後一個裡面

                        last_FromTo = item.From_To;
                        last_value = now_default_value;

                        #endregion 把現在的參數寄到最後一個裡面
                    }
                    else //現在的參數比上一個還要小
                    {
                        if (!errorKey.IsNullOrWhiteSpace()) //上一個是錯誤的,修改錯誤記錄資料
                        {
                            var hestory = overData[errorKey];
                            hestory.Add(item.From_To);
                            overData[errorKey] = hestory;
                        }
                        else //上一個是對的(這次錯誤需新增錯誤資料)
                        {
                            overData.Add(last_FromTo,
                                new List<string>() { last_FromTo, item.From_To }); //加入一筆歷史錯誤
                            errorKey = last_FromTo;//紀錄上一筆的FromTo為超過起始欄位
                        }
                        last_value = (last_value + now_default_value) / 2; //default 相加除以2
                    }
                }

                #endregion 找出錯誤的參數

                #region 組出DataTable 的欄位

                dt.Columns.Add("TM", typeof(object)); //第一欄固定為TM
                List<string> errorData = new List<string>(); //錯誤資料
                List<string> rowData = new List<string>(); //左邊行數欄位
                foreach (Moody_Tm_YYYY item in dbDatas) //第二次迴圈組 DataTable 欄位
                {
                    if (overData.ContainsKey(item.From_To)) //為起始錯誤
                    {
                        errorData = overData[item.From_To]; //把錯誤資料找出來
                    }
                    else if (errorData.Contains(item.From_To)) //為中間錯誤
                    {
                        //不做任何動作
                    }
                    else //無錯誤 (columns 加入原本 參數)
                    {
                        if (errorData.Any()) //上一筆是錯誤情形
                        {
                            string key =
                                string.Format("{0}_{1}",
                                errorData.First(),
                                errorData.Last()
                                );
                            dt.Columns.Add(key, typeof(object));
                            errorData = new List<string>();
                            rowData.Add(key);
                        }
                        dt.Columns.Add(item.From_To, typeof(object));
                        rowData.Add(item.From_To);
                    }
                }
                if (errorData.Any()) //此為最後一筆為錯誤時觸發
                {
                    string key =
                        string.Format("{0}_{1}",
                        errorData.First(),
                        errorData.Last()
                        );
                    dt.Columns.Add(key, typeof(double));
                    rowData.Add(key);
                }
                //最後兩欄固定為 WR & Default
                dt.Columns.Add("WR", typeof(string));
                dt.Columns.Add("Default", typeof(string));

                #endregion 組出DataTable 的欄位

                #region 組出資料

                List<string> columnsName = (from q in rowData select q).ToList();
                columnsName.AddRange(new List<string>() { "WR", "Default" });
                foreach (var item in rowData) //by 每一行
                {
                    if (item.IndexOf('_') > -1) //合併行需特別處理
                    {
                        List<string> err = overData[item.Split('_')[0]];

                        List<Moody_Tm_YYYY> dbs = dbDatas.Where(x => err.Contains(x.From_To)).ToList();
                        List<double> datas = new List<double>();
                        foreach (string cname in columnsName) //by 每一欄
                        {
                            if (cname.IndexOf('_') > -1) //合併欄
                            {
                                List<string> err2 = overData[cname.Split('_')[0]];
                                datas.Add((from y in dbs
                                           select (
                                           (from z in err2
                                            select getDbValueINColume(y, z))
                                           .Sum())).Sum() / (err.Count));
                            }
                            else
                            {
                                datas.Add((from x in dbs
                                           select getDbValueINColume(x, cname)).Sum()
                                            / err.Count);
                            }
                        }
                        List<object> o = new List<object>();
                        o.Add(item);
                        o.AddRange((from q in datas select q as object).ToList());
                        var row = dt.NewRow();
                        row.ItemArray = (o.ToArray());
                        dt.Rows.Add(row);
                    }
                    else //其他無合併行的只要單獨處理某特例欄位(合併)
                    {
                        string from_to = item;
                        List<double> datas = new List<double>();
                        Moody_Tm_YYYY db = dbDatas.Where(x => x.From_To == item).First();
                        foreach (string cname in columnsName) //by 每一欄
                        {
                            if (cname.IndexOf('_') > -1) //合併欄
                            {
                                List<string> err = overData[cname.Split('_')[0]];
                                //double avg = err.Select(x => getDbValueINColume(db, x)).Sum() / err.Count;
                                double avg = err.Select(x => getDbValueINColume(db, x)).Sum();
                                datas.Add(avg);
                            }
                            else //正常的
                            {
                                datas.Add(getDbValueINColume(db, cname));
                            }
                        }
                        List<object> o = new List<object>();
                        o.Add(item);
                        o.AddRange((from q in datas select q as object).ToList());
                        var row = dt.NewRow();
                        row.ItemArray = (o.ToArray());
                        dt.Rows.Add(row);
                    }
                }

                //加入 WT & Default 行
                List<string> WTArray = new List<string>() { "Baa1", "Baa2", "Baa3", "Ba1", "Ba2", "Ba3" };
                List<string> WTRow = new List<string>(); //WT要尋找的行的From/To
                foreach (var item in overData) //合併的資料紀錄
                {
                    if (WTArray.Intersect(item.Value).Any()) //找合併裡面符合的
                    {
                        WTRow.Add(item.Key);
                    }
                }
                WTRow.AddRange(rowData.Where(x => WTArray.Contains(x)));
                List<object> WRData = new List<object>();
                List<object> DefaultData = new List<object>();
                WRData.Add("WR");
                DefaultData.Add("Default");
                for (var i = 1; i < dt.Rows[0].ItemArray.Count(); i++) //i從1開始 from/to 那欄不用看
                {
                    double d = 0d;
                    for (var j = 0; j < dt.Rows.Count; j++)
                    {
                        if (WTRow.Contains(dt.Rows[j].ItemArray[0].ToString())) //符合排
                        {
                            d += Convert.ToDouble(dt.Rows[j][i]);
                        }
                    }
                    //WRData.Add(d * (dt.Columns[i].ToString().IndexOf("_") > -1 ?
                    //    overData[dt.Columns[i].ToString().Split('_')[0]].Count : 1)
                    //    / WTRow.Count); //WR多筆需*合併數                  
                    WRData.Add(d / WTRow.Count); //WR多筆不需合併數
                    if (i == (dt.Rows[0].ItemArray.Count() - 1))
                    {
                        DefaultData.Add(100d);
                    }
                    else
                    {
                        DefaultData.Add(0d);
                    }
                }
                var nrow = dt.NewRow();
                nrow.ItemArray = (WRData.ToArray());
                dt.Rows.Add(nrow);
                nrow = dt.NewRow();
                nrow.ItemArray = (DefaultData.ToArray());
                dt.Rows.Add(nrow);

                #endregion 組出資料
            }
            catch
            {
            }
            return new Tuple<DataTable, Dictionary<string, List<string>>>(dt, overData);
        }

        #endregion DB Moody_Tm_YYYY 組成 DataTable

        #region 抓DB的資料

        /// <summary>
        /// 抓DB的資料
        /// </summary>
        /// <param name="db">Moody_Tm_YYYY</param>
        /// <param name="cname">哪一欄位</param>
        /// <returns></returns>
        private double getDbValueINColume(Moody_Tm_YYYY db, string cname)
        {
            if (cname.Equals(A7_Type.Aaa.ToString()))
                return TypeTransfer.doubleNToDouble(db.Aaa);
            if (cname.Equals(A7_Type.Aa1.ToString()))
                return TypeTransfer.doubleNToDouble(db.Aa1);
            if (cname.Equals(A7_Type.Aa2.ToString()))
                return TypeTransfer.doubleNToDouble(db.Aa2);
            if (cname.Equals(A7_Type.Aa3.ToString()))
                return TypeTransfer.doubleNToDouble(db.Aa3);
            if (cname.Equals(A7_Type.A1.ToString()))
                return TypeTransfer.doubleNToDouble(db.A1);
            if (cname.Equals(A7_Type.A2.ToString()))
                return TypeTransfer.doubleNToDouble(db.A2);
            if (cname.Equals(A7_Type.A3.ToString()))
                return TypeTransfer.doubleNToDouble(db.A3);
            if (cname.Equals(A7_Type.Baa1.ToString()))
                return TypeTransfer.doubleNToDouble(db.Baa1);
            if (cname.Equals(A7_Type.Baa2.ToString()))
                return TypeTransfer.doubleNToDouble(db.Baa2);
            if (cname.Equals(A7_Type.Baa3.ToString()))
                return TypeTransfer.doubleNToDouble(db.Baa3);
            if (cname.Equals(A7_Type.Ba1.ToString()))
                return TypeTransfer.doubleNToDouble(db.Ba1);
            if (cname.Equals(A7_Type.Ba2.ToString()))
                return TypeTransfer.doubleNToDouble(db.Ba2);
            if (cname.Equals(A7_Type.Ba3.ToString()))
                return TypeTransfer.doubleNToDouble(db.Ba3);
            if (cname.Equals(A7_Type.B1.ToString()))
                return TypeTransfer.doubleNToDouble(db.B1);
            if (cname.Equals(A7_Type.B2.ToString()))
                return TypeTransfer.doubleNToDouble(db.B2);
            if (cname.Equals(A7_Type.B3.ToString()))
                return TypeTransfer.doubleNToDouble(db.B3);
            if (cname.Equals(A7_Type.Caa1.ToString()))
                return TypeTransfer.doubleNToDouble(db.Caa1);
            if (cname.Equals(A7_Type.Caa2.ToString()))
                return TypeTransfer.doubleNToDouble(db.Caa2);
            if (cname.Equals(A7_Type.Caa3.ToString()))
                return TypeTransfer.doubleNToDouble(db.Caa3);
            if (cname.Equals("Ca-C"))
                return TypeTransfer.doubleNToDouble(db.Ca_C);
            if (cname.Equals(A7_Type.WR.ToString()))
                return TypeTransfer.doubleNToDouble(db.WR);
            if (cname.Equals(A7_Type.Default.ToString()))
                return TypeTransfer.doubleNToDouble(db.Default_Value);
            return 0d;
        }

        #endregion 抓DB的資料

        #region A72 資料轉 A73

        /// <summary>
        /// A72 資料轉 A73
        /// </summary>
        /// <param name="dt">DataTable</param>
        /// <returns></returns>
        private DataTable FromA72GetA73(DataTable dt)
        {
            DataTable newData = new DataTable(); //要組新的 Table
            try
            {
                foreach (var itme in A73Array)
                {
                    newData.Columns.Add(itme, typeof(string)); //組 column
                }
                List<string>[] A73datas = new List<string>[A73Array.Count]; //需求的欄位資料
                for (int i = 0; i < A73Array.Count; i++)
                {
                    //取得需求的欄位資料
                    A73datas[i] = dt.AsEnumerable().Select(x => x.Field<string>(A73Array[i])).ToList();
                }
                if (A73datas.Any() && A73datas[0].Any()) //有資料
                {
                    for (int j = 0; j < A73datas[0].Count; j++) //原本datatable 的行數
                    {
                        List<string> o = new List<string>();
                        for (int k = 0; k < A73Array.Count; k++)
                        {
                            o.Add(A73datas[k][j]);
                        }
                        var row = newData.NewRow();
                        row.ItemArray = (o.ToArray());
                        newData.Rows.Add(row);
                    }
                }
            }
            catch
            {
            }
            return newData;
        }

        #endregion A72 資料轉 A73

        #region Create Table(DataTable 組 sql Create Table)

        private string DropCreateTable(string tableName, DataTable dt)
        {
            string sqlsc = string.Empty; //create table sql
            sqlsc += string.Format("{0} {1} {2} \n",
                 @" Begin Try drop table ",
                //@" drop table ",
                tableName,
                @" End Try Begin Catch End Catch "
                //@" "
                ); //有舊的table 刪除

            sqlsc += " CREATE TABLE " + tableName + "(";
            sqlsc += @" Id INT not null PRIMARY KEY , ";

            for (int i = 0; i < dt.Columns.Count; i++)
            {
                sqlsc += "\n [" + dt.Columns[i].ColumnName.Replace('-', '_')
                    .Replace("Default", "Default_Value") + "] ";
                if (0.Equals(i))
                {
                    sqlsc += " varchar(10) ";
                }
                else
                {
                    sqlsc += " float ";
                }
                sqlsc += " ,";
            }
            sqlsc = sqlsc.Substring(0, sqlsc.Length - 1) + "\n) ";
            return sqlsc;
        }

        /// <summary>
        /// 動態建Table sql 語法
        /// </summary>
        /// <param name="tableName"></param>
        /// <param name="dt"></param>
        /// <returns></returns>
        private string CreateA7Table(string tableName, DataTable dt)
        {
            //string sqlsc = string.Empty; //create table sql
            //sqlsc += string.Format("{0} {1} {2} \n",
            //     //@" Begin Try drop table ",
            //     @" drop table ",
            //    tableName,
            //    //@" End Try Begin Catch End Catch "
            //    @" "
            //    ); //有舊的table 刪除

            //sqlsc += " CREATE TABLE " + tableName + "(";
            //sqlsc += @" Id INT not null PRIMARY KEY , ";

            //for (int i = 0; i < dt.Columns.Count; i++)
            //{
            //    sqlsc += "\n [" + dt.Columns[i].ColumnName.Replace('-', '_')
            //        .Replace("Default", "Default_Value") + "] ";
            //    if (0.Equals(i))
            //    {
            //        sqlsc += " varchar(10) ";
            //    }
            //    else
            //    {
            //        sqlsc += " float ";
            //    }
            //    sqlsc += " ,";
            //}
            //sqlsc = sqlsc.Substring(0, sqlsc.Length - 1) + "\n) ";

            string sqlInsert = string.Empty; //insert sql
            int id = 1;
            for (var i = 0; i < dt.Rows.Count; i++) //每一行資料
            {
                string columnArray = string.Format(" {0} ,", "Id");
                string valueArray = string.Format(" {0} ,", id.ToString()); ;
                for (int j = 0; j < dt.Rows[i].ItemArray.Count(); j++)
                {
                    columnArray += string.Format(" {0} ,", dt.Columns[j].ToString()
                        .Replace('-', '_').Replace("Default", "Default_Value"));
                    if (0.Equals(j)) //第一筆是文字
                    {
                        valueArray += string.Format(" '{0}' ,", dt.Rows[i].ItemArray[j].ToString());
                    }
                    else
                    {
                        valueArray += string.Format(" {0} ,", dt.Rows[i].ItemArray[j].ToString());
                    }
                }
                sqlInsert += string.Format(" \n {0} {1} ({2}) {3} ({4}) ",
                             @" insert into ",
                             tableName,
                             columnArray.Substring(0, columnArray.Length - 1),
                             "values",
                             valueArray.Substring(0, valueArray.Length - 1)
                             );
                id += 1;
            }

            //return sqlsc + sqlInsert;
            return sqlInsert;
        }

        #endregion Create Table(DataTable 組 sql Create Table)

        #endregion Private Function
    }
}