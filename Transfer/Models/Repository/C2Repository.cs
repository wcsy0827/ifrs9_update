using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using Transfer.Models.Interface;
using Transfer.Utility;
using Transfer.ViewModels;
using static Transfer.Enum.Ref;

namespace Transfer.Models.Repository
{
    public class C2Repository : IC2Repository
    {
        public C2Repository()
        {
            this.common = new Common();
        }

        protected Common common
        {
            get;
            private set;
        }

        #region  getC13LagData
        public Tuple<bool, List<C13ViewModel>, List<jqGridColModel>, List<string>> getC13LagData(string tableName)
        {
            DataTable dt = getC23LagData(tableName);

            string lagNumber = "";

            int dataCount = dt.Rows.Count;

            if (dataCount == 0)
            {
                return new Tuple<bool, List<C13ViewModel>, List<jqGridColModel>, List<string>>(false, new List<C13ViewModel>(), new List<jqGridColModel>(), new List<string>());
            }

            string[] arrayTemp = dt.Columns[1].ColumnName.Split('_');
            lagNumber = arrayTemp[arrayTemp.Length - 1].Replace("L", "");
            string lagWord = $"_落後{lagNumber}期";

            int[] width = { 100, 200, 200, 200, 200 };
            var jqgridInfo = new C13ViewModel().TojqGridData(width);
            List<jqGridColModel> jqgridColModel = jqgridInfo.colModel;
            List<string> jqgridColNames = jqgridInfo.colNames;
            var tempName = "";
            List<string> removeColModel = new List<string>();
            List<string> removeColNames = new List<string>();
            for (int i = 0; i < jqgridColModel.Count; i++)
            {
                var IsTheSame = "N";

                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    tempName = dt.Columns[j].ColumnName;

                    arrayTemp = tempName.Split('_');

                    if (tempName != "Year_Quartly")
                    {
                        tempName = tempName.Replace("_" + arrayTemp[arrayTemp.Length - 1], "");
                    }

                    if (jqgridColModel[i].name == tempName)
                    {
                        IsTheSame = "Y";
                        break;
                    }
                }

                if (jqgridColModel[i].name != "Year_Quartly")
                {
                    jqgridColNames[i] = jqgridColNames[i] + lagWord;
                }

                if (IsTheSame == "N")
                {
                    removeColModel.Add(jqgridColModel[i].name);
                    removeColNames.Add(jqgridColNames[i]);
                }
            }

            for (int i = 0; i < removeColModel.Count; i++)
            {
                tempName = removeColModel[i];
                var oneColModel = jqgridColModel.Where(x => x.name == tempName).FirstOrDefault();
                if (oneColModel != null)
                {
                    jqgridColModel.Remove(oneColModel);
                }

                tempName = removeColNames[i];
                var oneColNames = jqgridColNames.Where(x => x == tempName).FirstOrDefault();
                if (oneColNames != null)
                {
                    jqgridColNames.Remove(oneColNames);
                }
            }

            List<C13ViewModel> listViewModel = new List<C13ViewModel>();

            for (int i = 0; i < dataCount; i++)
            {
                C13ViewModel oneViewModel = new C13ViewModel();

                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    tempName = dt.Columns[j].ColumnName;

                    var pro = oneViewModel.GetType().GetProperties()
                                          .Where(x => tempName.Contains(x.Name))
                                          .FirstOrDefault();
                    if (pro != null)
                    {
                        pro.SetValue(oneViewModel, dt.Rows[i][tempName].ToString());
                    }
                }

                listViewModel.Add(oneViewModel);
            }

            return new Tuple<bool, List<C13ViewModel>, List<jqGridColModel>, List<string>>(true, listViewModel, jqgridColModel, jqgridColNames);
        }
        #endregion

        #region getC23LagData
        private DataTable getC23LagData(string tableName)
        {
            DataTable dtTemp = new DataTable();
            DataTable dtResult = new DataTable();

            string cs = common.RemoveEntityFrameworkMetadata(string.Empty);
            using (var conn = new SqlConnection(cs))
            {
                string sql = $"SELECT * FROM sys.Tables WHERE name = '{tableName}'";
                var cmd = new SqlCommand();
                cmd.CommandText = sql;
                cmd.Connection = conn;
                conn.Open();
                dtTemp.Load(cmd.ExecuteReader());
                if (dtTemp.Rows.Count > 0)
                {
                    sql = $"SELECT * FROM {tableName}";
                    cmd.CommandText = sql;
                    dtResult.Load(cmd.ExecuteReader());
                    Extension.NlogSet(sql);
                }
            }

            return dtResult;
        }
        #endregion

        #region DownLoadC13Excel
        public MSGReturnModel DownLoadC13Excel(string path, List<C13ViewModel> dbDatas, List<jqGridColModel> jqgridColModel, List<string> jqgridColNames)
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            result.DESCRIPTION = Message_Type.download_Fail
                                 .GetDescription(Message_Type.not_Find_Any.GetDescription());

            if (dbDatas.Any())
            {
                DataTable datas = getC13ModelFromDb(dbDatas, jqgridColModel, jqgridColNames).Item1;

                result.DESCRIPTION = FileRelated.DataTableToExcel(datas, path, Excel_DownloadName.C13);
                result.RETURN_FLAG = string.IsNullOrWhiteSpace(result.DESCRIPTION);

                if (result.RETURN_FLAG)
                {
                    result.DESCRIPTION = Message_Type.download_Success.GetDescription();
                }
            }

            return result;
        }

        #endregion

        #region getC13ModelFromDb
        private Tuple<DataTable> getC13ModelFromDb(List<C13ViewModel> dbDatas, List<jqGridColModel> jqgridColModel, List<string> jqgridColNames)
        {
            DataTable dt = new DataTable();

            try
            {
                for (int i = 0; i < jqgridColNames.Count; i++)
                {
                    dt.Columns.Add(jqgridColNames[i], typeof(object));
                }

                foreach (C13ViewModel item in dbDatas)
                {
                    var nrow = dt.NewRow();

                    for (int i = 0; i < jqgridColNames.Count; i++)
                    {
                        var tempName = jqgridColModel[i].name;
                        var pro = item.GetType().GetProperties()
                                      .Where(x => x.Name == tempName)
                                      .FirstOrDefault();
                        if (pro != null)
                        {
                            nrow[jqgridColNames[i]] = (string)pro.GetValue(item, null); ;
                        }
                    }

                    dt.Rows.Add(nrow);
                }
            }
            catch
            {
            }

            return new Tuple<DataTable>(dt);
        }
        #endregion

        #region  getC13Column
        public Tuple<bool, List<C23ViewModel>> getC13Column()
        {
            var jqgridInfo = new C13ViewModel().TojqGridData();
            List<jqGridColModel> jqgridColModel = jqgridInfo.colModel;
            List<string> jqgridColNames = jqgridInfo.colNames;

            List<C23ViewModel> listC23 = new List<C23ViewModel>();

            for (int i = 0; i < jqgridColModel.Count; i++)
            {
                C23ViewModel oneC23 = new C23ViewModel();

                if (jqgridColModel[i].name == "Consumer_Price_Index" || jqgridColModel[i].name == "Unemployment_Rate")
                {
                    oneC23.Column_Name = jqgridColModel[i].name;
                    oneC23.Var_Name = jqgridColNames[i];

                    listC23.Add(oneC23);
                }
            }

            if (listC23.Any())
            {
                return new Tuple<bool, List<C23ViewModel>>(true, listC23);
            }

            return new Tuple<bool, List<C23ViewModel>>(false, new List<C23ViewModel>());
        }
        #endregion

        #region transferC13
        public MSGReturnModel transferC13(string lagNumber, string tableName, List<C23ViewModel> C23)
        {
            MSGReturnModel result = new MSGReturnModel();

            try
            {
                using (IFRS9DBEntities db = new IFRS9DBEntities())
                {
                    List<Econ_D_Var> query = db.Econ_D_Var.AsNoTracking().ToList();
                    if (query.Any() == false)
                    {
                        result.RETURN_FLAG = false;
                        result.DESCRIPTION = $"C13：Econ_D_Var 無資料";
                        return result;
                    }

                    int intLagNumber = int.Parse(lagNumber);

                    Econ_D_Var lastData = query[query.Count - 1];
                    int year = int.Parse(lastData.Year_Quartly.Substring(0, lastData.Year_Quartly.Length - 2));
                    int quartly = int.Parse(lastData.Year_Quartly.Substring(lastData.Year_Quartly.Length - 1,1));
                    for (int i = 1; i <= intLagNumber; i++)
                    {
                        quartly = quartly + 1;
                        
                        if (quartly > 4)
                        {
                            year = year + 1;
                            quartly = 1;
                        }

                        Econ_D_Var oneData = new Econ_D_Var();
                        oneData.Year_Quartly = year.ToString() + "Q" + quartly.ToString();
                        query.Add(oneData);
                    }

                    List<C13ViewModel> C13List = new List<C13ViewModel>();
                    
                    for (int i = 0; i < query.Count; i++)
                    {
                        C13ViewModel C13 = new C13ViewModel();

                        C13.Year_Quartly = query[i].Year_Quartly;

                        if (i >= intLagNumber)
                        {
                            C13.Consumer_Price_Index = query[i- intLagNumber].Consumer_Price_Index.ToString();
                            C13.Unemployment_Rate = query[i - intLagNumber].Unemployment_Rate.ToString();
                        }

                        C13List.Add(C13);
                    }

                    if (C13List.Count > 0)
                    {
                        CreateC13Table(tableName, C23, C13List);
                    }

                    result.RETURN_FLAG = true;
                    result.DESCRIPTION = "執行轉換成功";
                }

            }
            catch (Exception ex)
            {
                result.RETURN_FLAG = false;
                result.DESCRIPTION = ex.Message;
            }

            return result;
        }
        #endregion

        #region CreateC13Table
        private void CreateC13Table(string tableName, List<C23ViewModel> C23, List<C13ViewModel> data)
        {
            string columnName = "";

            string sqlsc = string.Empty;
            sqlsc += string.Format("{0} {1} {2}", " Begin Try drop table ", tableName, " End Try Begin Catch End Catch ");

            sqlsc += " CREATE TABLE " + tableName + "(";
            sqlsc += " Year_Quartly varchar(10) NOT NULL PRIMARY KEY,";
            for (int i = 0; i < C23.Count; i++)
            {
                sqlsc += C23[i].Column_Name + " float NULL,";

                columnName += C23[i].Column_Name + ",";
            }
            sqlsc = sqlsc.Substring(0, sqlsc.Length - 1) + "\n) ";
            columnName = columnName.Substring(0, columnName.Length - 1);

            string sqlInsert = string.Empty;

            for (var i = 0; i < data.Count; i++)
            {
                var dataValue = "";

                for (int j = 0; j < C23.Count; j++)
                {
                    var tempName = C23[j].Column_Name;
                    var pro = data[i].GetType().GetProperties()
                                     .Where(x => tempName.Contains(x.Name))
                                     .FirstOrDefault();
                    if (pro != null)
                    {
                        string tempValue = (string)pro.GetValue(data[i], null);

                        if (tempValue.IsNullOrWhiteSpace())
                        {
                            dataValue += "NULL,";
                        }
                        else
                        {
                            dataValue += tempValue + ",";
                        }
                    }
                }

                dataValue = dataValue.Substring(0, dataValue.Length - 1);

                sqlInsert += string.Format(" \n {0} {1} ({2}) {3} ({4});",
                                           " INSERT INTO ",
                                             tableName,
                                             "Year_Quartly," + columnName,
                                             "VALUES",
                                             "'" + data[i].Year_Quartly + "'," + dataValue
                                          );
            }

            string cs = common.RemoveEntityFrameworkMetadata(string.Empty);
            using (var conn = new SqlConnection(cs))
            {
                var cmd = new SqlCommand();
                cmd.CommandText = sqlsc;
                cmd.Connection = conn;
                conn.Open();
                cmd.ExecuteNonQuery();
                Extension.NlogSet(sqlsc);
                cmd.CommandText = sqlInsert;
                cmd.ExecuteNonQuery();
                Extension.NlogSet(sqlInsert);
                cmd.Dispose();
            }
        }
        #endregion

        #region  getC03LagData
        public Tuple<bool, List<C03ViewModel>, List<jqGridColModel>, List<string>> getC03LagData(string tableName)
        {
            DataTable dt = getC23LagData(tableName);

            string lagNumber = "";

            int dataCount = dt.Rows.Count;

            if (dataCount == 0)
            {
                return new Tuple<bool, List<C03ViewModel>, List<jqGridColModel>, List<string>>(false, new List<C03ViewModel>(), new List<jqGridColModel>(), new List<string>());
            }

            string[] arrayTemp = dt.Columns[1].ColumnName.Split('_');
            lagNumber = arrayTemp[arrayTemp.Length - 1].Replace("L", "");
            string lagWord = $"_落後{lagNumber}期";

            int[] width = { 100, 100, 100, 90, 100,
                            200,220,230,200,300,
                            230,230,230,230,230,
                            220,220,220,300,220,
                            400,250,300,300,250,
                            500,250,250,250,300,
                            300,300 };
            var jqgridInfo = new C03ViewModel().TojqGridData(width);
            List<jqGridColModel> jqgridColModel = jqgridInfo.colModel;
            List<string> jqgridColNames = jqgridInfo.colNames;
            var tempName = "";
            List<string> removeColModel = new List<string>();
            List<string> removeColNames = new List<string>();
            for (int i = 0; i < jqgridColModel.Count; i++)
            {
                var IsTheSame = "N";

                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    tempName = dt.Columns[j].ColumnName;

                    arrayTemp = tempName.Split('_');

                    if (tempName != "Year_Quartly")
                    {
                        tempName = tempName.Replace("_" + arrayTemp[arrayTemp.Length - 1], "");
                    }

                    if (jqgridColModel[i].name == tempName)
                    {
                        IsTheSame = "Y";
                        break;
                    }
                }

                if (jqgridColModel[i].name != "Year_Quartly")
                {
                    jqgridColNames[i] = jqgridColNames[i] + lagWord;
                }

                if (IsTheSame == "N")
                {
                    removeColModel.Add(jqgridColModel[i].name);
                    removeColNames.Add(jqgridColNames[i]);
                }
            }

            for (int i = 0; i < removeColModel.Count; i++)
            {
                tempName = removeColModel[i];
                var oneColModel = jqgridColModel.Where(x => x.name == tempName).FirstOrDefault();
                if (oneColModel != null)
                {
                    jqgridColModel.Remove(oneColModel);
                }

                tempName = removeColNames[i];
                var oneColNames = jqgridColNames.Where(x => x == tempName).FirstOrDefault();
                if (oneColNames != null)
                {
                    jqgridColNames.Remove(oneColNames);
                }
            }

            List<C03ViewModel> listViewModel = new List<C03ViewModel>();

            for (int i = 0; i < dataCount; i++)
            {
                C03ViewModel oneViewModel = new C03ViewModel();

                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    tempName = dt.Columns[j].ColumnName;

                    var pro = oneViewModel.GetType().GetProperties()
                                          .Where(x => tempName.Contains(x.Name))
                                          .FirstOrDefault();
                    if (pro != null)
                    {
                        pro.SetValue(oneViewModel, dt.Rows[i][tempName].ToString());
                    }
                }

                listViewModel.Add(oneViewModel);
            }

            return new Tuple<bool, List<C03ViewModel>, List<jqGridColModel>, List<string>>(true, listViewModel, jqgridColModel, jqgridColNames);
        }
        #endregion

        #region DownLoadC03Excel
        public MSGReturnModel DownLoadC03Excel(string path, List<C03ViewModel> dbDatas, List<jqGridColModel> jqgridColModel, List<string> jqgridColNames)
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            result.DESCRIPTION = Message_Type.download_Fail
                                 .GetDescription(Message_Type.not_Find_Any.GetDescription());
            if (dbDatas.Any())
            {
                DataTable datas = getC03ModelFromDb(dbDatas, jqgridColModel, jqgridColNames).Item1;

                result.DESCRIPTION = FileRelated.DataTableToExcel(datas, path, Excel_DownloadName.C03);
                result.RETURN_FLAG = string.IsNullOrWhiteSpace(result.DESCRIPTION);

                if (result.RETURN_FLAG)
                {
                    result.DESCRIPTION = Message_Type.download_Success.GetDescription();
                }
            }

            return result;
        }

        #endregion

        #region getC03ModelFromDb
        private Tuple<DataTable> getC03ModelFromDb(List<C03ViewModel> dbDatas, List<jqGridColModel> jqgridColModel, List<string> jqgridColNames)
        {
            DataTable dt = new DataTable();

            try
            {
                for (int i = 0; i < jqgridColNames.Count; i++)
                {
                    dt.Columns.Add(jqgridColNames[i], typeof(object));
                }

                foreach (C03ViewModel item in dbDatas)
                {
                    var nrow = dt.NewRow();

                    for (int i = 0; i < jqgridColNames.Count; i++)
                    {
                        var tempName = jqgridColModel[i].name;
                        var pro = item.GetType().GetProperties()
                                      .Where(x => x.Name == tempName)
                                      .FirstOrDefault();
                        if (pro != null)
                        {
                            nrow[jqgridColNames[i]] = (string)pro.GetValue(item, null); ;
                        }
                    }

                    dt.Rows.Add(nrow);
                }
            }
            catch
            {
            }

            return new Tuple<DataTable>(dt);
        }
        #endregion

        #region  getC03Column
        public Tuple<bool, List<C23ViewModel>> getC03Column()
        {
            var jqgridInfo = new C03ViewModel().TojqGridData();
            List<jqGridColModel> jqgridColModel = jqgridInfo.colModel;
            List<string> jqgridColNames = jqgridInfo.colNames;

            List<C23ViewModel> listC23 = new List<C23ViewModel>();

            for (int i = 0; i < jqgridColModel.Count; i++)
            {
                C23ViewModel oneC23 = new C23ViewModel();

                if (jqgridColModel[i].name != "Processing_Date"
                    && jqgridColModel[i].name != "Product_Code"
                    && jqgridColModel[i].name != "Data_ID"
                    && jqgridColModel[i].name != "Year_Quartly"
                    && jqgridColModel[i].name != "PD_Quartly")
                {
                    oneC23.Column_Name = jqgridColModel[i].name;
                    oneC23.Var_Name = jqgridColNames[i];

                    listC23.Add(oneC23);
                }
            }

            if (listC23.Any())
            {
                return new Tuple<bool, List<C23ViewModel>>(true, listC23);
            }

            return new Tuple<bool, List<C23ViewModel>>(false, new List<C23ViewModel>());
        }
        #endregion

        #region transferC03
        public MSGReturnModel transferC03(string lagNumber, string tableName, List<C23ViewModel> C23)
        {
            MSGReturnModel result = new MSGReturnModel();

            try
            {
                using (IFRS9DBEntities db = new IFRS9DBEntities())
                {
                    List<Econ_D_YYYYMMDD> query = db.Econ_D_YYYYMMDD.AsNoTracking().ToList();
                    if (query.Any() == false)
                    {
                        result.RETURN_FLAG = false;
                        result.DESCRIPTION = $"C03：Econ_D_YYYYMMDD 無資料";
                        return result;
                    }

                    int intLagNumber = int.Parse(lagNumber);

                    Econ_D_YYYYMMDD lastData = query[query.Count - 1];
                    int year = int.Parse(lastData.Year_Quartly.Substring(0, lastData.Year_Quartly.Length - 2));
                    int quartly = int.Parse(lastData.Year_Quartly.Substring(lastData.Year_Quartly.Length - 1, 1));
                    for (int i = 1; i <= intLagNumber; i++)
                    {
                        quartly = quartly + 1;

                        if (quartly > 4)
                        {
                            year = year + 1;
                            quartly = 1;
                        }

                        Econ_D_YYYYMMDD oneData = new Econ_D_YYYYMMDD();
                        oneData.Year_Quartly = year.ToString() + "Q" + quartly.ToString();
                        query.Add(oneData);
                    }

                    List<C03ViewModel> ViewModelList = new List<C03ViewModel>();

                    for (int i = 0; i < query.Count; i++)
                    {
                        C03ViewModel oneViewModel = new C03ViewModel();

                        oneViewModel.Year_Quartly = query[i].Year_Quartly;

                        if (i >= intLagNumber)
                        {
                            oneViewModel.TWSE_Index = query[i - intLagNumber].TWSE_Index.ToString();
                            oneViewModel.TWRGSARP_Index = query[i - intLagNumber].TWRGSARP_Index.ToString();
                            oneViewModel.TWGDPCON_Index = query[i - intLagNumber].TWGDPCON_Index.ToString();
                            oneViewModel.TWLFADJ_Index = query[i - intLagNumber].TWLFADJ_Index.ToString();
                            oneViewModel.TWCPI_Index = query[i - intLagNumber].TWCPI_Index.ToString();
                            oneViewModel.TWMSA1A_Index = query[i - intLagNumber].TWMSA1A_Index.ToString();
                            oneViewModel.TWMSA1B_Index = query[i - intLagNumber].TWMSA1B_Index.ToString();
                            oneViewModel.TWMSAM2_Index = query[i - intLagNumber].TWMSAM2_Index.ToString();
                            oneViewModel.GVTW10YR_Index = query[i - intLagNumber].GVTW10YR_Index.ToString();
                            oneViewModel.TWTRBAL_Index = query[i - intLagNumber].TWTRBAL_Index.ToString();
                            oneViewModel.TWTREXP_Index = query[i - intLagNumber].TWTREXP_Index.ToString();
                            oneViewModel.TWTRIMP_Index = query[i - intLagNumber].TWTRIMP_Index.ToString();
                            oneViewModel.TAREDSCD_Index = query[i - intLagNumber].TAREDSCD_Index.ToString();
                            oneViewModel.TWCILI_Index = query[i - intLagNumber].TWCILI_Index.ToString();
                            oneViewModel.TWBOPCUR_Index = query[i - intLagNumber].TWBOPCUR_Index.ToString();
                            oneViewModel.EHCATW_Index = query[i - intLagNumber].EHCATW_Index.ToString();
                            oneViewModel.TWINDPI_Index = query[i - intLagNumber].TWINDPI_Index.ToString();
                            oneViewModel.TWWPI_Index = query[i - intLagNumber].TWWPI_Index.ToString();
                            oneViewModel.TARSYOY_Index = query[i - intLagNumber].TARSYOY_Index.ToString();
                            oneViewModel.TWEOTTL_Index = query[i - intLagNumber].TWEOTTL_Index.ToString();
                            oneViewModel.SLDETIGT_Index = query[i - intLagNumber].SLDETIGT_Index.ToString();
                            oneViewModel.TWIRFE_Index = query[i - intLagNumber].TWIRFE_Index.ToString();
                            oneViewModel.SINYI_HOUSE_PRICE_index = query[i - intLagNumber].SINYI_HOUSE_PRICE_index.ToString();
                            oneViewModel.CATHAY_ESTATE_index = query[i - intLagNumber].CATHAY_ESTATE_index.ToString();
                            oneViewModel.Real_GDP2011 = query[i - intLagNumber].Real_GDP2011.ToString();
                            oneViewModel.MCCCTW_Index = query[i - intLagNumber].MCCCTW_Index.ToString();
                            oneViewModel.TRDR1T_Index = query[i - intLagNumber].TRDR1T_Index.ToString();
                        }

                        ViewModelList.Add(oneViewModel);
                    }

                    if (ViewModelList.Count > 0)
                    {
                        CreateC03Table(tableName, C23, ViewModelList);
                    }

                    result.RETURN_FLAG = true;
                    result.DESCRIPTION = "執行轉換成功";
                }

            }
            catch (Exception ex)
            {
                result.RETURN_FLAG = false;
                result.DESCRIPTION = ex.Message;
            }

            return result;
        }
        #endregion

        #region CreateC03Table
        private void CreateC03Table(string tableName, List<C23ViewModel> C23, List<C03ViewModel> data)
        {
            string columnName = "";

            string sqlsc = string.Empty;
            sqlsc += string.Format("{0} {1} {2}", " Begin Try drop table ", tableName, " End Try Begin Catch End Catch ");

            sqlsc += " CREATE TABLE " + tableName + "(";
            sqlsc += " Year_Quartly varchar(10) NOT NULL PRIMARY KEY,";
            for (int i = 0; i < C23.Count; i++)
            {
                sqlsc += C23[i].Column_Name + " float NULL,";

                columnName += C23[i].Column_Name + ",";
            }
            sqlsc = sqlsc.Substring(0, sqlsc.Length - 1) + "\n) ";
            columnName = columnName.Substring(0, columnName.Length - 1);

            string sqlInsert = string.Empty;

            for (var i = 0; i < data.Count; i++)
            {
                var dataValue = "";

                for (int j = 0; j < C23.Count; j++)
                {
                    var tempName = C23[j].Column_Name;
                    var pro = data[i].GetType().GetProperties()
                                     .Where(x => tempName.Contains(x.Name))
                                     .FirstOrDefault();
                    if (pro != null)
                    {
                        string tempValue = (string)pro.GetValue(data[i], null);

                        if (tempValue.IsNullOrWhiteSpace())
                        {
                            dataValue += "NULL,";
                        }
                        else
                        {
                            dataValue += tempValue + ",";
                        }
                    }
                }

                dataValue = dataValue.Substring(0, dataValue.Length - 1);

                sqlInsert += string.Format(" \n {0} {1} ({2}) {3} ({4});",
                                           " INSERT INTO ",
                                             tableName,
                                             "Year_Quartly," + columnName,
                                             "VALUES",
                                             "'" + data[i].Year_Quartly + "'," + dataValue
                                          );
            }

            string cs = common.RemoveEntityFrameworkMetadata(string.Empty);
            using (var conn = new SqlConnection(cs))
            {
                var cmd = new SqlCommand();
                cmd.CommandText = sqlsc;
                cmd.Connection = conn;
                conn.Open();
                cmd.ExecuteNonQuery();
                Extension.NlogSet(sqlsc);
                cmd.CommandText = sqlInsert;
                cmd.ExecuteNonQuery();
                Extension.NlogSet(sqlInsert);
                cmd.Dispose();
            }
        }
        #endregion

        #region  getC04LagData
        public Tuple<bool, List<C04ViewModel>, List<jqGridColModel>, List<string>> getC04LagData(string tableName)
        {
            DataTable dt = getC23LagData(tableName);

            string lagNumber = "";

            int dataCount = dt.Rows.Count;

            if (dataCount == 0)
            {
                return new Tuple<bool, List<C04ViewModel>, List<jqGridColModel>, List<string>>(false, new List<C04ViewModel>(), new List<jqGridColModel>(), new List<string>());
            }

            string[] arrayTemp = dt.Columns[1].ColumnName.Split('_');
            lagNumber = arrayTemp[arrayTemp.Length - 1].Replace("L", "");
            string lagWord = $"_落後{lagNumber}期";

            int[] width = {100, 100, 100, 90, 100,
                           220,220,220,220,220,220,220,220,220,220};
            var jqgridInfo = new C04ViewModel().TojqGridData(width);
            List<jqGridColModel> jqgridColModel = jqgridInfo.colModel;
            List<string> jqgridColNames = jqgridInfo.colNames;
            var tempName = "";
            List<string> removeColModel = new List<string>();
            List<string> removeColNames = new List<string>();
            for (int i = 0; i < jqgridColModel.Count; i++)
            {
                var IsTheSame = "N";

                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    tempName = dt.Columns[j].ColumnName;

                    arrayTemp = tempName.Split('_');

                    if (tempName != "Year_Quartly")
                    {
                        tempName = tempName.Replace("_" + arrayTemp[arrayTemp.Length - 1], "");
                    }

                    if (jqgridColModel[i].name == tempName)
                    {
                        IsTheSame = "Y";
                        break;
                    }
                }

                if (jqgridColModel[i].name != "Year_Quartly")
                {
                    jqgridColNames[i] = jqgridColNames[i] + lagWord;
                }

                if (IsTheSame == "N")
                {
                    removeColModel.Add(jqgridColModel[i].name);
                    removeColNames.Add(jqgridColNames[i]);
                }
            }

            for (int i = 0; i < removeColModel.Count; i++)
            {
                tempName = removeColModel[i];
                var oneColModel = jqgridColModel.Where(x => x.name == tempName).FirstOrDefault();
                if (oneColModel != null)
                {
                    jqgridColModel.Remove(oneColModel);
                }

                tempName = removeColNames[i];
                var oneColNames = jqgridColNames.Where(x => x == tempName).FirstOrDefault();
                if (oneColNames != null)
                {
                    jqgridColNames.Remove(oneColNames);
                }
            }

            List<C04ViewModel> listViewModel = new List<C04ViewModel>();

            for (int i = 0; i < dataCount; i++)
            {
                C04ViewModel oneViewModel = new C04ViewModel();

                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    tempName = dt.Columns[j].ColumnName;

                    var pro = oneViewModel.GetType().GetProperties()
                                          .Where(x => tempName.Contains(x.Name))
                                          .FirstOrDefault();
                    if (pro != null)
                    {
                        pro.SetValue(oneViewModel, dt.Rows[i][tempName].ToString());
                    }
                }

                listViewModel.Add(oneViewModel);
            }

            return new Tuple<bool, List<C04ViewModel>, List<jqGridColModel>, List<string>>(true, listViewModel, jqgridColModel, jqgridColNames);
        }
        #endregion

        #region DownLoadC04Excel
        public MSGReturnModel DownLoadC04Excel(string path, List<C04ViewModel> dbDatas, List<jqGridColModel> jqgridColModel, List<string> jqgridColNames)
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            result.DESCRIPTION = Message_Type.download_Fail
                                 .GetDescription(Message_Type.not_Find_Any.GetDescription());
            if (dbDatas.Any())
            {
                DataTable datas = getC04ModelFromDb(dbDatas, jqgridColModel, jqgridColNames).Item1;

                result.DESCRIPTION = FileRelated.DataTableToExcel(datas, path, Excel_DownloadName.C04);
                result.RETURN_FLAG = string.IsNullOrWhiteSpace(result.DESCRIPTION);

                if (result.RETURN_FLAG)
                {
                    result.DESCRIPTION = Message_Type.download_Success.GetDescription();
                }
            }

            return result;
        }
        #endregion

        #region getC04ModelFromDb
        private Tuple<DataTable> getC04ModelFromDb(List<C04ViewModel> dbDatas, List<jqGridColModel> jqgridColModel, List<string> jqgridColNames)
        {
            DataTable dt = new DataTable();

            try
            {
                for (int i = 0; i < jqgridColNames.Count; i++)
                {
                    dt.Columns.Add(jqgridColNames[i], typeof(object));
                }

                foreach (C04ViewModel item in dbDatas)
                {
                    var nrow = dt.NewRow();

                    for (int i = 0; i < jqgridColNames.Count; i++)
                    {
                        var tempName = jqgridColModel[i].name;
                        var pro = item.GetType().GetProperties()
                                      .Where(x => x.Name == tempName)
                                      .FirstOrDefault();
                        if (pro != null)
                        {
                            nrow[jqgridColNames[i]] = (string)pro.GetValue(item, null); ;
                        }
                    }

                    dt.Rows.Add(nrow);
                }
            }
            catch (Exception ex)
            {

            }

            return new Tuple<DataTable>(dt);
        }
        #endregion

        #region  getC04Column
        public Tuple<bool, List<C23ViewModel>> getC04Column()
        {
            var jqgridInfo = new C04ViewModel().TojqGridData();
            List<jqGridColModel> jqgridColModel = jqgridInfo.colModel;
            List<string> jqgridColNames = jqgridInfo.colNames;

            List<C23ViewModel> listC23 = new List<C23ViewModel>();

            for (int i = 0; i < jqgridColModel.Count; i++)
            {
                C23ViewModel oneC23 = new C23ViewModel();

                if (jqgridColModel[i].name != "Processing_Date"
                    && jqgridColModel[i].name != "Product_Code"
                    && jqgridColModel[i].name != "Data_ID"
                    && jqgridColModel[i].name != "Year_Quartly"
                    && jqgridColModel[i].name != "PD_Quartly")
                {
                    oneC23.Column_Name = jqgridColModel[i].name;
                    oneC23.Var_Name = jqgridColNames[i];

                    listC23.Add(oneC23);
                }
            }

            if (listC23.Any())
            {
                return new Tuple<bool, List<C23ViewModel>>(true, listC23);
            }

            return new Tuple<bool, List<C23ViewModel>>(false, new List<C23ViewModel>());
        }
        #endregion

        #region transferC04
        public MSGReturnModel transferC04(string lagNumber, string tableName, List<C23ViewModel> C23)
        {
            MSGReturnModel result = new MSGReturnModel();

            try
            {
                using (IFRS9DBEntities db = new IFRS9DBEntities())
                {
                    List<Econ_F_YYYYMMDD> query = db.Econ_F_YYYYMMDD.AsNoTracking().ToList();
                    if (query.Any() == false)
                    {
                        result.RETURN_FLAG = false;
                        result.DESCRIPTION = $"C04：Econ_F_YYYYMMDD 無資料";
                        return result;
                    }

                    int intLagNumber = int.Parse(lagNumber);

                    Econ_F_YYYYMMDD lastData = query[query.Count - 1];
                    int year = int.Parse(lastData.Year_Quartly.Substring(0, lastData.Year_Quartly.Length - 2));
                    int quartly = int.Parse(lastData.Year_Quartly.Substring(lastData.Year_Quartly.Length - 1, 1));
                    for (int i = 1; i <= intLagNumber; i++)
                    {
                        quartly = quartly + 1;

                        if (quartly > 4)
                        {
                            year = year + 1;
                            quartly = 1;
                        }

                        Econ_F_YYYYMMDD oneData = new Econ_F_YYYYMMDD();
                        oneData.Year_Quartly = year.ToString() + "Q" + quartly.ToString();
                        query.Add(oneData);
                    }

                    List<C04ViewModel> ViewModelList = new List<C04ViewModel>();

                    for (int i = 0; i < query.Count; i++)
                    {
                        C04ViewModel oneViewModel = new C04ViewModel();

                        oneViewModel.Year_Quartly = query[i].Year_Quartly;

                        if (i >= intLagNumber)
                        {
                            oneViewModel.CPI_INDX_Index = query[i - intLagNumber].CPI_INDX_Index.ToString();
                            oneViewModel.CACPSA_Index = query[i - intLagNumber].CACPSA_Index.ToString();
                            oneViewModel.KOCPI_Index = query[i - intLagNumber].KOCPI_Index.ToString();
                            oneViewModel.EACPI_Index = query[i - intLagNumber].EACPI_Index.ToString();
                            oneViewModel.GRCP2000_Index = query[i - intLagNumber].GRCP2000_Index.ToString();
                            oneViewModel.UKRPCHVJ_Index = query[i - intLagNumber].UKRPCHVJ_Index.ToString();
                            oneViewModel.NECPIND_Index = query[i - intLagNumber].NECPIND_Index.ToString();
                            oneViewModel.AUCCTOTS_Index = query[i - intLagNumber].AUCCTOTS_Index.ToString();
                            oneViewModel.JCPTSGEN_Index = query[i - intLagNumber].JCPTSGEN_Index.ToString();
                            oneViewModel.USGG3M_Index = query[i - intLagNumber].USGG3M_Index.ToString();
                            oneViewModel.GECU3M_Index = query[i - intLagNumber].GECU3M_Index.ToString();
                            oneViewModel.USGG10YR_Index = query[i - intLagNumber].USGG10YR_Index.ToString();
                            oneViewModel.GECU10YR_Index = query[i - intLagNumber].GECU10YR_Index.ToString();
                            oneViewModel.DEBPTOTL_Index = query[i - intLagNumber].DEBPTOTL_Index.ToString();
                            oneViewModel.NKY_Index = query[i - intLagNumber].NKY_Index.ToString();
                            oneViewModel.INDU_Index = query[i - intLagNumber].INDU_Index.ToString();
                            oneViewModel.CCMP_Index = query[i - intLagNumber].CCMP_Index.ToString();
                            oneViewModel.SPX_Index = query[i - intLagNumber].SPX_Index.ToString();
                            oneViewModel.DAX_Index = query[i - intLagNumber].DAX_Index.ToString();
                            oneViewModel.DJST_Index = query[i - intLagNumber].DJST_Index.ToString();
                            oneViewModel.USURTOT_Index = query[i - intLagNumber].USURTOT_Index.ToString();
                            oneViewModel.UMRTEMU_Index = query[i - intLagNumber].UMRTEMU_Index.ToString();
                            oneViewModel.JNUE_Index = query[i - intLagNumber].JNUE_Index.ToString();
                            oneViewModel.CANLXEMR_Index = query[i - intLagNumber].CANLXEMR_Index.ToString();
                            oneViewModel.GRUEPR_Index = query[i - intLagNumber].GRUEPR_Index.ToString();
                            oneViewModel.UKUEILOR_Index = query[i - intLagNumber].UKUEILOR_Index.ToString();
                            oneViewModel.NEUENTTR_Index = query[i - intLagNumber].NEUENTTR_Index.ToString();
                            oneViewModel.KOEAUERS_Index = query[i - intLagNumber].KOEAUERS_Index.ToString();
                            oneViewModel.AULFUNEM_Index = query[i - intLagNumber].AULFUNEM_Index.ToString();
                            oneViewModel.HPIMLEVL_Index = query[i - intLagNumber].HPIMLEVL_Index.ToString();
                            oneViewModel.GDP_CHWG_Index = query[i - intLagNumber].GDP_CHWG_Index.ToString();
                            oneViewModel.CGE9MP_Index = query[i - intLagNumber].CGE9MP_Index.ToString();
                            oneViewModel.EUGNEMU_Index = query[i - intLagNumber].EUGNEMU_Index.ToString();
                            oneViewModel.GRGDEGDP_Index = query[i - intLagNumber].GRGDEGDP_Index.ToString();
                            oneViewModel.UKGRABMI_Index = query[i - intLagNumber].UKGRABMI_Index.ToString();
                            oneViewModel.NEGDPESA_Index = query[i - intLagNumber].NEGDPESA_Index.ToString();
                            oneViewModel.JGDPGDP_Index = query[i - intLagNumber].JGDPGDP_Index.ToString();
                            oneViewModel.KOGDSTOT_Index = query[i - intLagNumber].KOGDSTOT_Index.ToString();
                            oneViewModel.AUNAGDP_Index = query[i - intLagNumber].AUNAGDP_Index.ToString();
                            oneViewModel.EUQDGEUR_Index = query[i - intLagNumber].EUQDGEUR_Index.ToString();
                            oneViewModel.USCABAL_Index = query[i - intLagNumber].USCABAL_Index.ToString();
                            oneViewModel.EHCAUS_Index = query[i - intLagNumber].EHCAUS_Index.ToString();
                            oneViewModel.GRCAEU_Index = query[i - intLagNumber].GRCAEU_Index.ToString();
                            oneViewModel.EHCADE_Index = query[i - intLagNumber].EHCADE_Index.ToString();
                            oneViewModel.EUSATOTN_Index = query[i - intLagNumber].EUSATOTN_Index.ToString();
                            oneViewModel.EUR003M_Index = query[i - intLagNumber].EUR003M_Index.ToString();
                            oneViewModel.EUORDEPO_Index = query[i - intLagNumber].EUORDEPO_Index.ToString();
                            oneViewModel.USDR1T_CMPN_Curncy = query[i - intLagNumber].USDR1T_CMPN_Curncy.ToString();
                            oneViewModel.FDTR_Index = query[i - intLagNumber].FDTR_Index.ToString();
                            oneViewModel.SXEE_Index = query[i - intLagNumber].SXEE_Index.ToString();
                            oneViewModel.NG1_Comdty = query[i - intLagNumber].NG1_Comdty.ToString();
                            oneViewModel.CO1_COMDTY = query[i - intLagNumber].CO1_COMDTY.ToString();
                            oneViewModel.CL1_Comdty = query[i - intLagNumber].CL1_Comdty.ToString();
                            oneViewModel.XOI_Index = query[i - intLagNumber].XOI_Index.ToString();
                            oneViewModel.EURUSD_Curncy = query[i - intLagNumber].EURUSD_Curncy.ToString();
                            oneViewModel.USDCNY_Curncy = query[i - intLagNumber].USDCNY_Curncy.ToString();
                            oneViewModel.USDJPY_Curncy = query[i - intLagNumber].USDJPY_Curncy.ToString();
                            oneViewModel.USDKRW_Curncy = query[i - intLagNumber].USDKRW_Curncy.ToString();
                            oneViewModel.USDTWD_Curncy = query[i - intLagNumber].USDTWD_Curncy.ToString();
                            oneViewModel.EURTWD_Curncy = query[i - intLagNumber].EURTWD_Curncy.ToString();
                            oneViewModel.CNYTWD_Curncy = query[i - intLagNumber].CNYTWD_Curncy.ToString();
                            oneViewModel.JPYTWD_Curncy = query[i - intLagNumber].JPYTWD_Curncy.ToString();
                            oneViewModel.KRWTWD_Curncy = query[i - intLagNumber].KRWTWD_Curncy.ToString();
                            oneViewModel.EURR002W_Index = query[i - intLagNumber].EURR002W_Index.ToString();
                            oneViewModel.NAPMPMI_Index = query[i - intLagNumber].NAPMPMI_Index.ToString();
                            oneViewModel.NAPMNMI_Index = query[i - intLagNumber].NAPMNMI_Index.ToString();
                            oneViewModel.NAPMNEWO_Index = query[i - intLagNumber].NAPMNEWO_Index.ToString();
                            oneViewModel.OUTFGAF_Index = query[i - intLagNumber].OUTFGAF_Index.ToString();
                            oneViewModel.LEI_WKIJ_Index = query[i - intLagNumber].LEI_WKIJ_Index.ToString();
                            oneViewModel.LEI_LCI_Index = query[i - intLagNumber].LEI_LCI_Index.ToString();
                            oneViewModel.USHBMIDX_Index = query[i - intLagNumber].USHBMIDX_Index.ToString();
                            oneViewModel.ARDIMONY_Index = query[i - intLagNumber].ARDIMONY_Index.ToString();
                            oneViewModel.M1_Index = query[i - intLagNumber].M1_Index.ToString();
                            oneViewModel.ECMAM1_Index = query[i - intLagNumber].ECMAM1_Index.ToString();
                            oneViewModel.M2_Index = query[i - intLagNumber].M2_Index.ToString();
                            oneViewModel.ECMAM2_Index = query[i - intLagNumber].ECMAM2_Index.ToString();
                            oneViewModel.CEEREU_Index = query[i - intLagNumber].CEEREU_Index.ToString();
                            oneViewModel.DXY_Curncy = query[i - intLagNumber].DXY_Curncy.ToString();
                            oneViewModel.EUBCI_Index = query[i - intLagNumber].EUBCI_Index.ToString();
                            oneViewModel.OEDEA044_Index = query[i - intLagNumber].OEDEA044_Index.ToString();
                            oneViewModel.EUESEMU_Index = query[i - intLagNumber].EUESEMU_Index.ToString();
                            oneViewModel.SNTEEUGX_Index = query[i - intLagNumber].SNTEEUGX_Index.ToString();
                            oneViewModel.XTSBEZ_Index = query[i - intLagNumber].XTSBEZ_Index.ToString();
                            oneViewModel.GRTBALE_Index = query[i - intLagNumber].GRTBALE_Index.ToString();
                            oneViewModel.EBLS11NC_Index = query[i - intLagNumber].EBLS11NC_Index.ToString();
                            oneViewModel.MXWO0MT_Index = query[i - intLagNumber].MXWO0MT_Index.ToString();
                            oneViewModel.S5MATRX_Index = query[i - intLagNumber].S5MATRX_Index.ToString();
                            oneViewModel.XAU_Curncy = query[i - intLagNumber].XAU_Curncy.ToString();
                            oneViewModel.USTBTOT_Index = query[i - intLagNumber].USTBTOT_Index.ToString();
                            oneViewModel.ADP_CHNG_Index = query[i - intLagNumber].ADP_CHNG_Index.ToString();
                            oneViewModel.NFP_TCH_Index = query[i - intLagNumber].NFP_TCH_Index.ToString();
                            oneViewModel.CNSTTMOM_Index = query[i - intLagNumber].CNSTTMOM_Index.ToString();
                            oneViewModel.TMNOCHNG_Index = query[i - intLagNumber].TMNOCHNG_Index.ToString();
                            oneViewModel.DGNOCHNG_Index = query[i - intLagNumber].DGNOCHNG_Index.ToString();
                            oneViewModel.IP_CHNG_Index = query[i - intLagNumber].IP_CHNG_Index.ToString();
                            oneViewModel.NHSPSTOT_Index = query[i - intLagNumber].NHSPSTOT_Index.ToString();
                            oneViewModel.NHCHSTCH_Index = query[i - intLagNumber].NHCHSTCH_Index.ToString();
                            oneViewModel.NHSPATOT_Index = query[i - intLagNumber].NHSPATOT_Index.ToString();
                            oneViewModel.NHCHATCH_Index = query[i - intLagNumber].NHCHATCH_Index.ToString();
                            oneViewModel.LEI_CHNG_Index = query[i - intLagNumber].LEI_CHNG_Index.ToString();
                            oneViewModel.ETSLTOTL_Index = query[i - intLagNumber].ETSLTOTL_Index.ToString();
                            oneViewModel.ETSLMOM_Index = query[i - intLagNumber].ETSLMOM_Index.ToString();
                            oneViewModel.NHSLTOT_Index = query[i - intLagNumber].NHSLTOT_Index.ToString();
                            oneViewModel.NHSLCHNG_Index = query[i - intLagNumber].NHSLCHNG_Index.ToString();
                            oneViewModel.CONSSENT_Index = query[i - intLagNumber].CONSSENT_Index.ToString();
                            oneViewModel.RSTAMOM_Index = query[i - intLagNumber].RSTAMOM_Index.ToString();
                            oneViewModel.EUPPEMUM_Index = query[i - intLagNumber].EUPPEMUM_Index.ToString();
                            oneViewModel.EUCCEMU_Index = query[i - intLagNumber].EUCCEMU_Index.ToString();
                            oneViewModel.CPMINDX_Index = query[i - intLagNumber].CPMINDX_Index.ToString();
                            oneViewModel.MPMICNMA_Index = query[i - intLagNumber].MPMICNMA_Index.ToString();
                            oneViewModel.CNFRBAL_Index = query[i - intLagNumber].CNFRBAL_Index.ToString();
                            oneViewModel.CNCPIYOY_Index = query[i - intLagNumber].CNCPIYOY_Index.ToString();
                            oneViewModel.CHEFTYOY_Index = query[i - intLagNumber].CHEFTYOY_Index.ToString();
                            oneViewModel.CNMS2YOY_Index = query[i - intLagNumber].CNMS2YOY_Index.ToString();
                            oneViewModel.CNFRIMPY_Index = query[i - intLagNumber].CNFRIMPY_Index.ToString();
                            oneViewModel.CNFREXPY_Index = query[i - intLagNumber].CNFREXPY_Index.ToString();
                            oneViewModel.GRFIDEBT_Index = query[i - intLagNumber].GRFIDEBT_Index.ToString();
                        }

                        ViewModelList.Add(oneViewModel);
                    }

                    if (ViewModelList.Count > 0)
                    {
                        CreateC04Table(tableName, C23, ViewModelList);
                    }

                    result.RETURN_FLAG = true;
                    result.DESCRIPTION = "執行轉換成功";
                }

            }
            catch (Exception ex)
            {
                result.RETURN_FLAG = false;
                result.DESCRIPTION = ex.Message;
            }

            return result;
        }
        #endregion

        #region CreateC04Table
        private void CreateC04Table(string tableName, List<C23ViewModel> C23, List<C04ViewModel> data)
        {
            string columnName = "";

            string sqlsc = string.Empty;
            sqlsc += string.Format("{0} {1} {2}", " Begin Try drop table ", tableName, " End Try Begin Catch End Catch ");

            sqlsc += " CREATE TABLE " + tableName + "(";
            sqlsc += " Year_Quartly varchar(10) NOT NULL PRIMARY KEY,";
            for (int i = 0; i < C23.Count; i++)
            {
                sqlsc += C23[i].Column_Name + " float NULL,";

                columnName += C23[i].Column_Name + ",";
            }
            sqlsc = sqlsc.Substring(0, sqlsc.Length - 1) + "\n) ";
            columnName = columnName.Substring(0, columnName.Length - 1);

            string sqlInsert = string.Empty;

            for (var i = 0; i < data.Count; i++)
            {
                var dataValue = "";

                for (int j = 0; j < C23.Count; j++)
                {
                    var tempName = C23[j].Column_Name;
                    var pro = data[i].GetType().GetProperties()
                                     .Where(x => tempName.Contains(x.Name))
                                     .FirstOrDefault();
                    if (pro != null)
                    {
                        string tempValue = (string)pro.GetValue(data[i], null);

                        if (tempValue.IsNullOrWhiteSpace())
                        {
                            dataValue += "NULL,";
                        }
                        else
                        {
                            dataValue += tempValue + ",";
                        }
                    }
                }

                dataValue = dataValue.Substring(0, dataValue.Length - 1);

                sqlInsert += string.Format(" \n {0} {1} ({2}) {3} ({4});",
                                           " INSERT INTO ",
                                             tableName,
                                             "Year_Quartly," + columnName,
                                             "VALUES",
                                             "'" + data[i].Year_Quartly + "'," + dataValue
                                          );
            }

            string cs = common.RemoveEntityFrameworkMetadata(string.Empty);
            using (var conn = new SqlConnection(cs))
            {
                var cmd = new SqlCommand();
                cmd.CommandText = sqlsc;
                cmd.Connection = conn;
                conn.Open();
                cmd.ExecuteNonQuery();
                Extension.NlogSet(sqlsc);
                cmd.CommandText = sqlInsert;
                cmd.ExecuteNonQuery();
                Extension.NlogSet(sqlInsert);
                cmd.Dispose();
            }
        }
        #endregion
    }
 }