using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Web.Mvc;
using Transfer.Infrastructure;
using Transfer.Models.Interface;
using Transfer.Models.Repository;
using Transfer.Utility;
using Transfer.ViewModels;
using static Transfer.Enum.Ref;

namespace Transfer.Controllers
{
    [LogAuth]
    public class A8Controller : CommonController
    {
        private IA8Repository A8Repository;

        public A8Controller()
        {
            this.A8Repository = new A8Repository();
        }


        /// <summary>
        /// A8(查詢(A81.A82.A83))
        /// </summary>
        /// <returns></returns>
        [UserAuth]
        public ActionResult Detail()
        {
            return View();
        }

        /// <summary>
        /// 前端抓資料時呼叫
        /// </summary>
        /// <param name="type"></param>
        /// <returns></returns>
        [BrowserEvent("查詢A8系列資料")]
        [HttpPost]
        public JsonResult GetData(string type)
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            result.DESCRIPTION = Message_Type.not_Find_Any.GetDescription(type);
            try
            {
                switch (type)
                {
                    case "A81": //抓Moody_Monthly_PD_Info(A81)資料
                        var A81Data = A8Repository.GetA81();
                        result.RETURN_FLAG = A81Data.Item1;
                        result.Datas = Json(A81Data.Item2);
                        break;

                    case "A82"://抓Moody_Quartly_PD_Info(A82)資料
                        var A82Data = A8Repository.GetA82();
                        result.RETURN_FLAG = A82Data.Item1;
                        result.Datas = Json(A82Data.Item2);
                        break;

                    case "A83"://抓Moody_Predit_PD_Info(A83)資料
                        var A83Data = A8Repository.GetA83();
                        result.RETURN_FLAG = A83Data.Item1;
                        result.Datas = Json(A83Data.Item2);
                        break;
                }
            }
            catch (Exception ex)
            {
                result.RETURN_FLAG = false;
                result.DESCRIPTION = Message_Type.not_Find_Any.GetDescription(type, ex.Message);
            }
            return Json(result);
        }

        /// <summary>
        /// A8(上傳檔案)
        /// </summary>
        /// <returns></returns>
        [UserAuth]
        public ActionResult Index()
        {
            return View();
        }

        /// <summary>
        /// 轉檔把Excel 資料存到 DB
        /// </summary>
        /// <returns></returns>
        [BrowserEvent("Exhibit 10Excel檔存到DB")]
        [HttpPost]
        public JsonResult Transfer()
        {
            MSGReturnModel result = new MSGReturnModel();
            try
            {
                #region 抓Excel檔案 轉成 model

                // Excel 檔案位置
                DateTime startTime = DateTime.Now;
                string projectFile = Server.MapPath("~/" + SetFile.FileUploads);
                string fileName = string.Empty;
                if (Cache.IsSet(CacheList.A81ExcelName))
                    fileName = (string)Cache.Get(CacheList.A81ExcelName);  //從Cache 抓資料
                if (fileName.IsNullOrWhiteSpace())
                {
                    result.RETURN_FLAG = false;
                    result.DESCRIPTION = Message_Type.time_Out.GetDescription();
                }
                string path = Path.Combine(projectFile, fileName);
                FileStream stream = System.IO.File.Open(path, FileMode.Open, FileAccess.Read);

                string pathType = path.Split('.')[1]; //抓副檔名
                List<Exhibit10Model> dataModel = A8Repository.getExcel(pathType, stream); //Excel轉成 Exhibit10Model

                #endregion 抓Excel檔案 轉成 model

                #region txtlog 檔案名稱

                string txtpath = SetFile.A81TransferTxtLog; //預設txt名稱
                string configTxtName = ConfigurationManager.AppSettings["txtLogA8Name"];
                if (!string.IsNullOrWhiteSpace(configTxtName))
                    txtpath = configTxtName; //有設定webConfig且不為空就取代

                #endregion txtlog 檔案名稱

                #region save Moody_Monthly_PD_Info(A81)

                MSGReturnModel resultA81 = A8Repository.SaveA8(Table_Type.A81.ToString(), dataModel); //save to DB
                bool A81Log = CommonFunction.saveLog(Table_Type.A81, fileName, SetFile.ProgramName,
                    resultA81.RETURN_FLAG, Debt_Type.B.ToString(), startTime, DateTime.Now, AccountController.CurrentUserInfo.Name,0); //寫sql Log
                TxtLog.txtLog(Table_Type.A81, resultA81.RETURN_FLAG, startTime, txtLocation(txtpath)); //寫txt Log

                #endregion save Moody_Monthly_PD_Info(A81)

                #region save Moody_Quartly_PD_Info(A82)

                MSGReturnModel resultA82 = A8Repository.SaveA8(Table_Type.A82.ToString(), dataModel); //save to DB
                bool A82Log = CommonFunction.saveLog(Table_Type.A82, fileName, SetFile.ProgramName,
                    resultA82.RETURN_FLAG, Debt_Type.B.ToString(), startTime, DateTime.Now, AccountController.CurrentUserInfo.Name,0); //寫sql Log
                TxtLog.txtLog(Table_Type.A82, resultA82.RETURN_FLAG, startTime, txtLocation(txtpath)); //寫txt Log

                #endregion save Moody_Quartly_PD_Info(A82)

                #region save Moody_Predit_PD_Info(A83)

                MSGReturnModel resultA83 = A8Repository.SaveA8(Table_Type.A83.ToString(), dataModel); //save to DB
                bool A83Log = CommonFunction.saveLog(Table_Type.A83, fileName, SetFile.ProgramName,
                    resultA83.RETURN_FLAG, Debt_Type.B.ToString(), startTime, DateTime.Now, AccountController.CurrentUserInfo.Name,0); //寫sql Log
                TxtLog.txtLog(Table_Type.A83, resultA83.RETURN_FLAG, startTime, txtLocation(txtpath)); //寫txt Log

                #endregion save Moody_Predit_PD_Info(A83)

                result.RETURN_FLAG = resultA81.RETURN_FLAG &&
                                     resultA82.RETURN_FLAG &&
                                     resultA83.RETURN_FLAG;
                result.DESCRIPTION = Message_Type.save_Success.GetDescription(
                    string.Format("{0},{1},{2}",
                    Table_Type.A81.ToString(),
                    Table_Type.A82.ToString(),
                    Table_Type.A83.ToString()
                    ));

                if (!result.RETURN_FLAG)
                {
                    List<string> errs = new List<string>();
                    if (!resultA81.RETURN_FLAG)
                        errs.Add(resultA81.DESCRIPTION);
                    if (!resultA82.RETURN_FLAG)
                        errs.Add(resultA82.DESCRIPTION);
                    if (!resultA83.RETURN_FLAG)
                        errs.Add(resultA83.DESCRIPTION);

                    result.DESCRIPTION = string.Join("\n", errs);
                }
            }
            catch (Exception ex)
            {
                result.RETURN_FLAG = false;
                result.DESCRIPTION = Message_Type.save_Fail.GetDescription(null, ex.Message);
            }
            return Json(result);
        }

        /// <summary>
        /// 選擇檔案後點選資料上傳觸發
        /// </summary>
        /// <param name="FileModel"></param>
        /// <returns>MSGReturnModel</returns>
        [BrowserEvent("上傳Exhibit 10Excel檔案")]
        [HttpPost]
        public JsonResult Upload(ValidateFiles FileModel)
        {
            MSGReturnModel result = new MSGReturnModel();
            try
            {
                #region 前端無傳送檔案進來

                if (FileModel.File == null)
                {
                    result.RETURN_FLAG = false;
                    result.DESCRIPTION = Message_Type.upload_Not_Find.GetDescription();
                    return Json(result);
                }

                #endregion 前端無傳送檔案進來

                #region 前端檔案大小不服或不為Excel檔案(驗證)

                if (!ModelState.IsValid)
                {
                    result.RETURN_FLAG = false;
                    result.DESCRIPTION = Message_Type.excel_Validate.GetDescription();
                    return Json(result);
                }

                #endregion 前端檔案大小不服或不為Excel檔案(驗證)

                #region 上傳檔案

                string pathType = Path.GetExtension(FileModel.File.FileName)
                      .Substring(1); //上傳的檔案類型

                var fileName = string.Format("{0}.{1}",
                      Excel_UploadName.A81.GetDescription(),
                      pathType); //固定轉成此名稱

                Cache.Invalidate(CacheList.A81ExcelName); //清除 Cache
                Cache.Set(CacheList.A81ExcelName, fileName); //把資料存到 Cache

                #region 檢查是否有FileUploads資料夾,如果沒有就新增 並加入 excel 檔案

                string projectFile = Server.MapPath("~/" + SetFile.FileUploads); //專案資料夾
                string path = Path.Combine(projectFile, fileName);

                FileRelated.createFile(projectFile); //檢查是否有FileUploads資料夾,如果沒有就新增

                //呼叫上傳檔案 function
                result = FileRelated.FileUpLoadinPath(path, FileModel.File);

                if (!result.RETURN_FLAG)
                    return Json(result);

                #endregion 檢查是否有FileUploads資料夾,如果沒有就新增 並加入 excel 檔案

                #region 讀取Excel資料 使用ExcelDataReader 並且組成 json

                var stream = FileModel.File.InputStream;
                List<Exhibit10Model> dataModel = A8Repository.getExcel(pathType, stream);
                if (dataModel.Count > 0)
                {
                    result.RETURN_FLAG = true;
                    result.Datas = Json(dataModel); //給JqGrid 顯示
                }
                else
                {
                    result.RETURN_FLAG = false;
                    result.DESCRIPTION = Message_Type.data_Not_Compare.GetDescription();
                }

                #endregion 讀取Excel資料 使用ExcelDataReader 並且組成 json

                #endregion 上傳檔案
            }
            catch (Exception ex)
            {
                result.RETURN_FLAG = false;
                result.DESCRIPTION = Message_Type.upload_Fail.GetDescription(null, ex.Message);
            }
            return Json(result);
        }
    }
}