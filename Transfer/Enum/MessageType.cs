using System.ComponentModel;

namespace Transfer.Enum
{
    public partial class Ref
    {
        /// <summary>
        /// 回傳訊息格式統一
        /// </summary>
        public enum Message_Type
        {

            //增加C10覆蓋轉檔資料前提醒，
            /// <summary>
            /// 已經有上傳資料，請問要進行上傳並覆蓋資料?
            /// </summary>
            [Description("已經有上傳資料，請問要進行上傳並覆蓋資料?")]
            Uplaod_data_Overwrite_File,

            //190222 增加接不到A59資料錯誤訊息
            /// <summary>
            /// 沒有找到寶碩原始信評資料
            /// </summary>
            [Description("沒有找到寶碩原始信評資料!")]
            not_Find_CounterPartyCreditRating,
            /// <summary>
            /// 接取寶碩原始信評失敗
            /// </summary>
            [Description("接取寶碩原始信評失敗!")]
            Find_CounterPartyCreditRating_Error,
            /// <summary>
            /// 寶碩信評合併A59失敗
            /// </summary>
            [Description("寶碩信評合併A59失敗!")]
            Join_CounterPartyCreditRating_Error,
            /// <summary>
            /// 寶碩原始信評合併A59成功
            /// </summary>
            [Description("接取寶碩原始信評成功!")]
            Join_CounterPartyCreditRating_Success,
            //190222 增加找無對應缺漏資料訊息
            /// <summary>
            /// 找無對應缺漏資料
            /// </summary>
            [Description("沒有找到對應缺漏資料!")]
            Not_Find_Relating_Missing,



            /// <summary>
            /// 資料已經儲存過了
            /// </summary>
            [Description("資料已經儲存過了!")]
            already_Save,

            /// <summary>
            /// 已執行減損階段最終確認無法再執行
            /// </summary>
            [Description("已執行減損階段最終確認無法再執行")]
            already_Save_D54,

            /// <summary>
            /// 資料已異動(請重新查詢)
            /// </summary>
            [Description("資料已異動(請重新查詢)")]
            already_Change,

            /// <summary>
            /// 儲存成功
            /// </summary>
            [Description("儲存成功!")]
            save_Success,

            /// <summary>
            /// 儲存失敗
            /// </summary>
            [Description("儲存失敗!")]
            save_Fail,

            /// <summary>
            /// 修改成功
            /// </summary>
            [Description("修改成功!")]
            update_Success,

            /// <summary>
            /// 修改失敗
            /// </summary>
            [Description("修改失敗!")]
            update_Fail,

            /// <summary>
            /// 刪除成功
            /// </summary>
            [Description("刪除成功!")]
            delete_Success,

            /// <summary>
            /// 刪除失敗
            /// </summary>
            [Description("刪除失敗!")]
            delete_Fail,

            /// <summary>
            /// 查無資料
            /// </summary>
            [Description("查無資料!")]
            not_Find_Data,

            /// <summary>
            /// 沒有找到資料
            /// </summary>
            [Description("沒有找到資料!")]
            not_Find_Any,

            /// <summary>
            /// 無更新資料
            /// </summary>
            [Description("無更新資料!")]
            not_Find_Update_Data,

            /// <summary>
            /// 沒有找到搜尋的資料
            /// </summary>
            [Description("沒有找到搜尋的資料!")]
            query_Not_Find,

            /// <summary>
            /// 下載成功
            /// </summary>
            [Description("下載成功!")]
            download_Success,

            /// <summary>
            /// 下載失敗
            /// </summary>
            [Description("下載失敗!")]
            download_Fail,

            /// <summary>
            /// 上傳成功
            /// </summary>
            [Description("上傳成功!")]
            upload_Success,

            /// <summary>
            /// 上傳失敗
            /// </summary>
            [Description("上傳失敗!")]
            upload_Fail,

            /// <summary>
            /// 請選擇上傳檔案
            /// </summary>
            [Description("請選擇上傳檔案!")]
            upload_Not_Find,

            /// <summary>
            /// 請確認檔案為Excel檔案或無資料!
            /// </summary>
            [Description("請確認檔案為Excel檔案或無資料!")]
            excel_Validate,

            /// <summary>
            /// 無比對到資料!
            /// </summary>
            [Description("無比對到資料!")]
            data_Not_Compare,

            /// <summary>
            /// 傳入參數錯誤!
            /// </summary>
            [Description("傳入參數錯誤!")]
            parameter_Error,

            /// <summary>
            /// 時間停滯太久請重新上一動作!
            /// </summary>
            [Description("時間停滯太久請重新上一動作!")]
            time_Out,

            /// <summary>
            /// 轉檔驗證失敗
            /// </summary>
            [Description("轉檔驗證失敗!")]
            transferError,

            /// <summary>
            /// 執行減損計算失敗
            /// </summary>
            [Description("執行減損計算失敗!")]
            kriskError,

            /// <summary>
            /// 已新增資料等待呈送覆核!
            /// </summary>
            [Description("已新增資料等待呈送覆核!")]
            insert_Success_Wait_Audit,

            /// <summary>
            /// 已更新資料等待呈送覆核!
            /// </summary>
            [Description("已更新資料等待呈送覆核!")]
            update_Success_Wait_Audit,

            /// <summary>
            /// 已為失效狀態無須再設定失效!
            /// </summary>
            [Description("已為失效狀態無須再設定失效!")]
            already_IsActiveN,

            /// <summary>
            /// 資料尚未完成刪除，等待呈送覆核!
            /// </summary>
            [Description("資料尚未完成刪除，等待呈送覆核!")]
            delete_Success_Wait_Audit,

            /// <summary>
            /// 呈送複核成功
            /// </summary>
            [Description("呈送複核成功!")]
            send_To_Audit_Success,

            /// <summary>
            /// 呈送複核失敗
            /// </summary>
            [Description("呈送複核失敗!")]
            send_To_Audit_Fail,

            /// <summary>
            /// 複核成功
            /// </summary>
            [Description("複核成功!")]
            Audit_Success,

            /// <summary>
            /// 複核失敗
            /// </summary>
            [Description("複核失敗!")]
            Audit_Fail,

            /// <summary>
            /// 駁回成功
            /// </summary>
            [Description("駁回成功!")]
            Reject_Success,

            /// <summary>
            /// 無呈送複核權限
            /// </summary>
            [Description("無呈送複核權限!")]
            none_Send_Audit_Authority,

            /// <summary>
            /// 無複核權限
            /// </summary>
            [Description("無複核權限!")]
            none_Audit_Authority,

            /// <summary>
            /// 請輸入正確的帳號或密碼!
            /// </summary>
            [Description("請輸入正確的帳號或密碼!")]
            login_Fail,

            /// <summary>
            /// 帳號已失效!
            /// </summary>
            [Description("帳號已失效!")]
            login_Effective_Fail,

            /// <summary>
            /// 帳號登入中!
            /// </summary>
            [Description("帳號登入中!")]
            login_Flag_Fail,

            /// <summary>
            /// 請輸入正確的驗證碼!
            /// </summary>
            [Description("請輸入正確的驗證碼!")]
            login_Captcha_Fail,

            /// <summary>
            /// 推送評估成功!
            /// </summary>
            [Description("推送評估成功!")]
            push_Assessor_Success,

            /// <summary>
            /// 前置表檢核錯誤!
            /// </summary>
            [Description("前置表檢核錯誤!")]
            table_Check_Fail,

            /// <summary>
            /// 檢核錯誤!
            /// </summary>
            [Description("檢核錯誤!")]
            Check_Fail,

            /// <summary>
            /// 資料重複輸入!，請檢查資料:
            /// </summary>
            [Description("資料重複輸入!，請檢查資料:")]
            Check_Redundancy,


            /// <summary>
            /// 無法辨認資料為本金還是利息，請檢查資料格式，錯誤資料:
            /// </summary>
            [Description("無法辨認資料為本金還是利息，請檢查資料格式，錯誤資料:")]
            Check_TypeFail,



            /// <summary>
            /// 資料的本金或利息有缺漏，缺漏資料
            /// </summary>
            [Description("資料的本金或利息有缺漏，缺漏資料:")]
            DataLostFail,

            /// <summary>
            ///補上傳之資料皆無對應之部位
            /// </summary>
            [Description("補上傳之資料皆無對應之部位")]
            DataNoneCorrespond,

            #region PGE需求延伸，D54Insert中投會要求客製訊息
            /// <summary>
            /// 選擇的報導日無A41部位上傳
            /// </summary>
            [Description("選擇的報導日無A41部位上傳")]
            A41NotFind,
            /// <summary>
            /// 選擇的報導日無C10風控上傳excel計算檔
            /// </summary>
            [Description("選擇的報導日無C10風控上傳excel計算檔")]
            C10NotFind,
            /// <summary>
            /// A41註記部位與C10上傳檔部位不符
            /// </summary>
            [Description("此版本A41註記為”N”毋須進行評估部位與C10 風控上傳檔部位不相符。 請聯繫投資風控確認C10上傳檔")]
            A41C10NotMatch
            #endregion





        }
    }
}