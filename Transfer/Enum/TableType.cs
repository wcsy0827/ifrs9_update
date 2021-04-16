using System.ComponentModel;
using Transfer.Infrastructure;
using Transfer.ViewModels;

namespace Transfer.Enum
{
    public partial class Ref
    {
        /// <summary>
        /// 所有Table
        /// </summary>
        public enum Table_Type
        {
            //190222 自動補檔增加TableType，只為了顯示訊息，無實際Table
            /// <summary>
            /// A59自動由寶碩補入信評
            /// </summary>
            [Description("A59Apex")]
            [Name("寶碩DB接入信評")]
            A59Apex,
            /// <summary>
            /// A59轉入A57A58
            /// </summary>
            [Description("A59Transfer")]
            [Name("A59信評補登轉入A57A58")]
            A59Trans,

            /// <summary>
            /// 國內總體經濟變數
            /// </summary>
            [Description("Econ_Domestic")]
            [Name("國內總體經濟變數")]
            A07,

            /// <summary>
            /// 債券明細檔
            /// </summary>
            [Description("Bond_Account_Info")]
            [Name("債券明細檔")]
            A41,

            /// <summary>
            /// 國庫券月結資料檔
            /// </summary>
            [Description("Treasury_Securities_Info")]
            [Name("國庫券月結資料檔")]
            A42,

            /// <summary>
            /// 換券收息調整資料檔
            /// </summary>
            [Description("Bond_ISIN_Changed_IntRevise")]
            [Name("換券收息調整資料檔")]
            A44_2,     

            /// <summary>
            /// 產業別資訊檔
            /// </summary>
            [Description("Bond_Category_Info")]
            [Name("產業別資訊檔")]
            A45,

            /// <summary>
            /// 固收提供的CEIC資料
            /// </summary>
            [Commucation(typeof(A46ViewModel))]
            [Description("Fixed_Income_CEIC_Info")]
            [Name("固收提供的CEIC資料")]
            A46,

            /// <summary>
            /// 固收提供的Moody資料
            /// </summary>
            [Commucation(typeof(A47ViewModel))]
            [Description("Fixed_Income_Moody_Info")]
            [Name("固收提供的Moody資料")]
            A47,

            /// <summary>
            /// 固收提供的阿布達比資料
            /// </summary>
            [Commucation(typeof(A48ViewModel))]
            [Description("Fixed_Income_AbuDhabi_Info")]
            [Name("固收提供的阿布達比資料")]
            A48,

            /// <summary>
            /// 財報揭露會計帳值資料
            /// </summary>
            [Commucation(typeof(A49ViewModel))]
            [Description("Bond_Accounting_EL")]
            [Name("財報揭露會計帳值資料")]
            A49,

            /// <summary>
            /// 信評主標尺對應檔_Moody
            /// </summary>
            [Description("Grade_Moody_Info")]
            [Name("信評主標尺對應檔_Moody")]
            A51,

            /// <summary>
            /// 信評主標尺對應檔_Moody
            /// </summary>
            [Description("Grade_Mapping_Info")]
            [Name("信評主標尺對應檔_其他")]
            A52,

            /// <summary>
            /// 外部信評資料檔
            /// </summary>
            [Description("Rating_Info")]
            [Name("外部信評資料檔")]
            A53,

            /// <summary>
            /// 信評設殊值補正設定檔
            /// </summary>
            [Commucation(typeof(A56ViewModel))]
            [Description("Rating_Update_Info")]
            [Name("信評設殊值補正設定檔")]
            A56,

            /// <summary>
            /// 債券信評檔_歷史檔
            /// </summary>
            [Description("Bond_Rating_Info")]
            [Name("債券信評檔_歷史檔")]
            A57,

            /// <summary>
            /// 債券信評檔_整理檔
            /// </summary>
            [Description("Bond_Rating_Summary")]
            [Name("債券信評檔_整理檔")]
            A58,

            /// <summary>
            /// 債券信評補登批次檔
            /// </summary>
            [Description("Bond_Rating_Update_Info")]
            [Name("債券信評補登批次檔")]
            A59,

            /// <summary>
            /// 回收率資料檔_Moody
            /// </summary>
            [Description("Moody_Recovery_Info")]
            [Name("回收率資料檔_Moody")]
            A61,

            /// <summary>
            /// 違約損失資料檔_歷史資料
            /// </summary>
            [Description("Moody_LGD_Info")]
            [Name("違約損失資料檔_歷史資料")]
            A62,

            /// <summary>
            /// 違約損失資料檔_歷史資料
            /// </summary>
            [Description("Moody_LGD_Info")]
            [Name("違約損失資料檔_歷史資料")]
            A63,

            /// <summary>
            /// 轉移矩陣資料檔_Moody
            /// </summary>
            [Description("Moody_Tm_YYYY")]
            [Name("轉移矩陣資料檔_Moody")]
            A71,

            /// <summary>
            /// 轉移矩陣資料檔_調整後
            /// </summary>
            [Description("Tm_Adjust_YYYY")]
            [Name("轉移矩陣資料檔_調整後")]
            A72,

            /// <summary>
            /// 等級違約率矩陣
            /// </summary>
            [Description("GM_YYYY")]
            [Name("等級違約率矩陣")]
            A73,

            /// <summary>
            /// 月違約機率資料檔_Moody
            /// </summary>
            [Description("Moody_Monthly_PD_Info")]
            [Name("月違約機率資料檔_Moody")]
            A81,

            /// <summary>
            /// 季違約機率資料檔_Moody
            /// </summary>
            [Description("Moody_Quartly_PD_Info")]
            [Name("季違約機率資料檔_Moody")]
            A82,

            /// <summary>
            /// 機率預測資料_Moody
            /// </summary>
            [Description("Moody_Predit_PD_Info")]
            [Name("機率預測資料_Moody")]
            A83,

            /// <summary>
            /// 國外總體經濟變數
            /// </summary>
            [Description("Econ_Foreign")]
            [Name("國外總體經濟變數")]
            A84,

            /// <summary>
            /// 主權債測試指標_Ticker
            /// </summary>
            [Description("Gov_Info_Ticker")]
            [Name("主權債測試指標_Ticker")]
            A94,

            /// <summary>
            /// 債券種類與減損階段評估類別檔
            /// </summary>
            [Description("Bond_Ticker_Info")]
            [Name("債券種類與減損階段評估類別檔")]
            A95,

            /// <summary>
            /// 產業別對應Ticker補登檔
            /// </summary>
            [Commucation(typeof(A95_1ViewModel))]
            [Description("Assessment_Sub_Kind_Ticker")]
            [Name("產業別對應Ticker補登檔")]
            A95_1,

            /// <summary>
            /// 信用利差資料
            /// </summary>
            [Commucation(typeof(A96ViewModel))]
            [Description("Bond_Spread_Info")]
            [Name("信用利差資料")]
            A96,

            /// <summary>
            /// 信用利差最後交易日資料
            /// </summary>
            [Commucation(typeof(A96TradeViewModel))]
            [Description("Bond_Spread_Trade_Info")]
            [Name("信用利差最後交易日資料")]
            A96_Trade,

            /// <summary>
            /// 帳戶主檔
            /// </summary>
            [Description("IFRS9_Main")]
            [Name("帳戶主檔")]
            B01,

            /// <summary>
            /// 減損計算輸入資料
            /// </summary>
            [Description("EL_Data_In")]
            [Name("減損計算輸入資料")]
            C01,

            /// <summary>
            /// 減損計算輸入資料(香港)
            /// </summary>
            [Description("EL_Data_In")]
            [Name("減損計算輸入資料(香港)")]
            C01_HK,

            /// <summary>
            /// 減損計算輸入資料(越南)
            /// </summary>
            [Description("EL_Data_In")]
            [Name("減損計算輸入資料(越南)")]
            C01_VN,

            /// <summary>
            /// 信評歷史資料
            /// </summary>
            [Description("Rating_History")]
            [Name("信評歷史資料")]
            C02,

            /// <summary>
            /// 前瞻性國內模型資料
            /// </summary>
            [Description("Econ_D_YYYYMMDD ")]
            [Name("前瞻性國內模型資料")]
            C03,

            /// <summary>
            /// 前瞻性國外總經資料
            /// </summary>
            [Description("Econ_Foreign")]
            [Name("前瞻性國外總經資料")]
            C04,

            /// <summary>
            /// 減損計算輸出資料
            /// </summary>
            [Description("EL_Data_Out")]
            [Name("減損計算輸出資料")]
            C07,

            /// <summary>
            /// 減損計算輸入資料流程內修正
            /// </summary>
            [Description("EL_Data_In_update")]
            [Name("減損計算輸入資料流程內修正")]
            C09,

            /// <summary>
            /// 減損計算補上傳要件評估資料
            /// </summary>
            [Description("Bond_Account_AssessmentCheck")]
            [Name("減損計算補上傳要件評估資料")]
            C10,


            /// <summary>
            /// 報表需求資訊
            /// </summary>
            [Description("IFRS9_Bond_Report")]
            [Name("報表需求資訊")]
            D52,

            /// <summary>
            /// SMF對應表
            /// </summary>
            [Commucation(typeof(D53ViewModel))]
            [Description("SMF_Info")]
            [Name("SMF對應表")]
            D53,

            /// <summary>
            /// 最終減損金額
            /// </summary>
            [Description("IFRS9_EL")]
            [Name("最終減損金額")]
            D54,

            /// <summary>
            /// 信評優先選擇參數檔
            /// </summary>
            [Description("Bond_Rating_Parm")]
            [Name("信評優先選擇參數檔")]
            D60,

            /// <summary>
            /// 量化評估需求資訊檔
            /// </summary>
            [Description("Bond_Quantitative_Resource")]
            [Name("量化評估需求資訊檔")]
            D63,

            /// <summary>
            /// 量化評估結果檔
            /// </summary>
            [Description("Bond_Quantitative_Result")]
            [Name("量化評估結果檔")]
            D64,

            /// <summary>
            /// 質化評估資訊檔
            /// </summary>
            [Description("Bond_Qualitative_Assessment")]
            [Name("質化評估資訊檔")]
            D65,

            /// <summary>
            /// 質化評估結果檔
            /// </summary>
            [Description("Bond_Qualitative_Assessment_Result")]
            [Name("質化評估結果檔")]
            D66,

            /// <summary>
            /// 風控覆核專區
            /// </summary>
            [Description("Bond_RiskReview_Result_File")]
            [Name("風控覆核專區")]
            D6RiskReview,            

            /// <summary>
            /// SMF分群設定檔
            /// </summary>
            [Commucation(typeof(D72ViewModel))]
            [Description("SMF_Group")]
            [Name("SMF分群設定檔")]
            D72,

            /// <summary>
            /// 排程查核結果紀錄檔
            /// </summary>
            [Commucation(typeof(D73ViewModel))]
            [Description("Scheduling_Report")]
            [Name("排程查核結果紀錄檔")]
            D73,

            /// <summary>
            /// 通知設定
            /// </summary>
            [Commucation(typeof(D74ViewModel))]
            [Description("Notice_Info")]
            [Name("通知設定")]
            D74,

            /// <summary>
            /// 郵件設定
            /// </summary>
            [Commucation(typeof(D74_1ViewModel))]
            [Description("Mail_Info")]
            [Name("郵件設定")]
            D74_1,

            /// <summary>
            /// 風控覆核專區上傳檔
            /// </summary>
            [Description("Bond_Risk_Control_Result_File")]
            [Name("風控覆核專區上傳檔")]
            D75,

            /// <summary>
            /// 檢查轉檔完成資料表
            /// </summary>
            [Description("Transfer_CheckTable")]
            [Name("檢查轉檔完成資料表")]
            CheckTable,

            /// <summary>
            /// 帳戶資料表
            /// </summary>
            [Description("IFRS9_User")]
            [Name("帳戶資料表")]
            UserTable,

            /// <summary>
            /// 主權債測試指標_年收集
            /// </summary>
            [Description("Gov_Info_Yearly")]
            [Name("國內總體經濟變數")]
            A91,

            /// <summary>
            /// 主權債測試指標_季收集
            /// </summary>
            [Description("Gov_Info_Quartly")]
            [Name("主權債測試指標_季收集")]
            A92,

            /// <summary>
            /// 主權債測試指標_月收集
            /// </summary>
            [Description("Gov_Info_Monthly")]
            [Name("主權債測試指標_月收集")]
            A93,

            /// <summary>
            /// 信用利差資料
            /// </summary>
            [Description("Bond_Spread_Info")]
            [Name("信用利差資料")]
            A961,

            /// <summary>
            /// 信用利差資料
            /// </summary>
            [Description("Bond_Spread_Info")]
            [Name("信用利差資料")]
            A962,

            /// <summary>
            /// IFRS9Log
            /// </summary>
            [Description("IFRS9Log")]
            [Name("執行紀綠")]
            IFRS9Log,

        }
    }
}