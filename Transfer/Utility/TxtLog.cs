using System;
using System.IO;
using static Transfer.Enum.Ref;

namespace Transfer.Utility
{
    public static class TxtLog
    {
        #region save txtlog

        /// <summary>
        /// 寫入 Txt Log
        /// </summary>
        /// <param name="tableName">table名</param>
        /// <param name="falg">成功或失敗</param>
        /// <param name="start">開始時間</param>
        /// <param name="folderPath">檔案路徑</param>
        public static void txtLog(Table_Type tableName, bool falg, DateTime start, string folderPath)
        {
            try
            {
                string txtData = string.Empty;
                try //試著抓取舊資料
                {
                    txtData = File.ReadAllText(folderPath, System.Text.Encoding.Default);
                }
                catch { }
                string txt = string.Format("{0}({1})_{2}_{3}",
                             tableName.GetDescription(),
                             tableName.ToString(),
                             start.ToString("yyyyMMddHHmmss"),
                             falg ? "Y" : "N");
                writeTxt(folderPath, txt, txtData);
                //if (!string.IsNullOrWhiteSpace(txtData)) //有舊資料就換行寫入下一筆
                //{
                //    txtData += string.Format("\r\n{0}", txt);
                //}
                //else //沒有就直接寫入
                //{
                //    txtData = txt;
                //}
                //using (FileStream fs = new FileStream(folderPath, FileMode.Create, FileAccess.Write))
                //{
                //    StreamWriter sw = new StreamWriter(fs, System.Text.Encoding.Default);
                //    sw.Write(txtData); //存檔
                //    sw.Close();
                //}
            }
            catch 
            {
            }
        }

        #endregion save txtlog

        #region readLog

        /// <summary>
        /// 讀取txt 資料
        /// </summary>
        /// <param name="folderPath"></param>
        /// <returns></returns>
        public static string getLog(string folderPath)
        {
            string result = string.Empty;
            try
            {
                result = File.ReadAllText(folderPath, System.Text.Encoding.Default);
            }
            catch { }
            return result;
        }

        #endregion

        #region SendMailLog

        /// <summary>
        /// 寫入 寄信txt (量化評估使用)
        /// </summary>
        /// <param name="msg"></param>
        /// <param name="folderPath"></param>
        public static void senMailLog(string msg, string folderPath)
        {
            try {
                string txtData = string.Empty;
                try //試著抓取舊資料
                {
                    txtData = File.ReadAllText(folderPath, System.Text.Encoding.Default);
                }
                catch { }
                writeTxt(folderPath, msg, txtData,true);
            }
            catch  {

            }
        }

        /// <summary>
        /// 寫入 txt
        /// </summary>
        /// <param name="folderPath"></param>
        /// <param name="newtTxt"></param>
        /// <param name="oldTxt"></param>
        private static void writeTxt(string folderPath, string newtTxt, string oldTxt = null,bool flag = false)
        {
            if(!folderPath.IsNullOrWhiteSpace() && !newtTxt.IsNullOrWhiteSpace())
                using (FileStream fs = new FileStream(folderPath, FileMode.Create, FileAccess.Write))
                {
                    var txtData = string.Empty;
                    if (!oldTxt.IsNullOrWhiteSpace())
                    {
                        txtData = oldTxt;
                        if (!flag)
                            txtData += string.Format("\r\n{0}", newtTxt);
                        else
                            txtData = (string.Format("{0}\r\n{1}", newtTxt, txtData));
                    }
                    else
                        txtData = newtTxt;
                    StreamWriter sw = new StreamWriter(fs, System.Text.Encoding.Default);
                    sw.Write(txtData); //存檔
                    sw.Close();
                }
        } 
        #endregion


    }
}