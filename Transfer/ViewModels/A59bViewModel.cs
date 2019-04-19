using System.ComponentModel;
using Transfer.Infrastructure;

namespace Transfer.ViewModels
{
    public class A59bViewModel
    {
        /// <summary>
        /// 帳戶編號
        /// </summary>
        [Description("帳戶編號")]
        public string Reference_Nbr { get; set; }

        /// <summary>
        /// 報導日
        /// </summary>
        [Description("報導日")]
        public string Report_Date { get; set; }

        /// <summary>
        /// 資料版本
        /// </summary>
        [Description("資料版本")]
        public string Version { get; set; }

        /// <summary>
        /// 債券編號
        /// </summary>
        [Description("債券編號")]
        public string Bond_Number { get; set; }

        /// <summary>
        /// Lots
        /// </summary>
        [Description("Lots")]
        public string Lots { get; set; }

        /// <summary>
        /// 債券購入(認列)日期
        /// </summary>
        [Description("債券購入(認列)日期")]
        public string Origination_Date { get; set; }

        /// <summary>
        /// Portfolio英文
        /// </summary>
        [Description("Portfolio英文")]
        public string Portfolio_Name { get; set; }

        /// <summary>
        /// 債券產品別(揭露使用)
        /// </summary>
        [Description("債券產品別(揭露使用)")]
        public string SMF { get; set; }

        /// <summary>
        /// Rating_Type
        /// </summary>
        public string Rating_Type { get; set; }

        /// <summary>
        /// Issuer
        /// </summary>
        public string Issuer { get; set; }

        /// <summary>
        /// Security_Ticker
        /// </summary>
        public string Security_Ticker { get; set; }

        /// <summary>
        /// RATING_AS_OF_DATE_OVERRIDE
        /// </summary>
        public string RATING_AS_OF_DATE_OVERRIDE { get; set; }

        /// <summary>
        /// ISSUER_EQUITY_TICKER
        /// </summary>
        public string ISSUER_EQUITY_TICKER { get; set; }

        /// <summary>
        /// GUARANTOR_NAME
        /// </summary>
        public string GUARANTOR_NAME { get; set; }

        /// <summary>
        /// GUARANTOR_EQY_TICKER
        /// </summary>
        public string GUARANTOR_EQY_TICKER { get; set; }

        /// <summary>
        /// 債項_標普評等 (債項\ sp\國外)
        /// </summary>
        [SetDateFiled("SP_EFF_DT")]
        public string RTG_SP { get; set; }

        /// <summary>
        /// 債項_標普評等日期
        /// </summary>
        public string SP_EFF_DT { get; set; }

        /// <summary>
        /// 債項_TRC 評等 (債項\ CW\國內)
        /// </summary>
        [SetDateFiled("TRC_EFF_DT")]
        public string RTG_TRC { get; set; }

        /// <summary>
        /// 債項_TRC 評等日期
        /// </summary>
        public string TRC_EFF_DT { get; set; }

        /// <summary>
        /// 債項_穆迪評等 (債項\ moody\國外)
        /// </summary>
        [SetDateFiled("MOODY_EFF_DT")]
        public string RTG_MOODY { get; set; }

        /// <summary>
        /// 債項_穆迪評等日期
        /// </summary>
        public string MOODY_EFF_DT { get; set; }

        /// <summary>
        /// 債項_惠譽評等 (債項\ Fitch\國外)
        /// </summary>
        [SetDateFiled("FITCH_EFF_DT")]
        public string RTG_FITCH { get; set; }

        /// <summary>
        /// 債項_惠譽評等日期
        /// </summary>
        public string FITCH_EFF_DT { get; set; }

        /// <summary>
        /// 債項_惠譽國內評等 (債項\ Fitch(twn)\國內)
        /// </summary>
        [SetDateFiled("RTG_FITCH_NATIONAL_DT")]
        public string RTG_FITCH_NATIONAL { get; set; }

        /// <summary>
        /// 債項_惠譽國內評等日期
        /// </summary>
        public string RTG_FITCH_NATIONAL_DT { get; set; }

        /// <summary>
        /// 標普長期外幣發行人信用評等 (發行人\ sp\國外)
        /// </summary>
        [SetDateFiled("RTG_SP_LT_FC_ISS_CRED_RTG_DT")]
        public string RTG_SP_LT_FC_ISSUER_CREDIT { get; set; }

        /// <summary>
        /// 標普長期外幣發行人信用評等日期
        /// </summary>
        public string RTG_SP_LT_FC_ISS_CRED_RTG_DT { get; set; }

        /// <summary>
        /// 標普本國貨幣長期發行人信用評等 (發行人\ sp\國外)
        /// </summary>
        [SetDateFiled("RTG_SP_LT_LC_ISS_CRED_RTG_DT")]
        public string RTG_SP_LT_LC_ISSUER_CREDIT { get; set; }

        /// <summary>
        /// 標普本國貨幣長期發行人信用評等日期
        /// </summary>
        public string RTG_SP_LT_LC_ISS_CRED_RTG_DT { get; set; }

        /// <summary>
        /// 穆迪發行人評等 (發行人\ moody\國外)
        /// </summary>
        [SetDateFiled("RTG_MDY_ISSUER_RTG_DT")]
        public string RTG_MDY_ISSUER { get; set; }

        /// <summary>
        /// 穆迪發行人評等日期
        /// </summary>
        public string RTG_MDY_ISSUER_RTG_DT { get; set; }

        /// <summary>
        /// 發行人_穆迪長期評等 (發行人\ moody\國外)
        /// </summary>
        [SetDateFiled("RTG_MOODY_LONG_TERM_DATE")]
        public string RTG_MOODY_LONG_TERM { get; set; }

        /// <summary>
        /// 發行人_穆迪長期評等日期
        /// </summary>
        public string RTG_MOODY_LONG_TERM_DATE { get; set; }

        /// <summary>
        /// 發行人_穆迪優先無擔保債務評等 (發行人\ moody\國外)
        /// </summary>
        [SetDateFiled("RTG_MDY_SEN_UNSEC_RTG_DT")]
        public string RTG_MDY_SEN_UNSECURED_DEBT { get; set; }

        /// <summary>
        /// 發行人_穆迪優先無擔保債務評等_日期
        /// </summary>
        public string RTG_MDY_SEN_UNSEC_RTG_DT { get; set; }

        /// <summary>
        /// 穆迪外幣發行人評等 (發行人\ moody\國外)
        /// </summary>
        [SetDateFiled("RTG_MDY_FC_CURR_ISSUER_RTG_DT")]
        public string RTG_MDY_FC_CURR_ISSUER_RATING { get; set; }

        /// <summary>
        /// 穆迪外幣發行人評等日期
        /// </summary>
        public string RTG_MDY_FC_CURR_ISSUER_RTG_DT { get; set; }

        //2017/12/25 A59 欄位不顯示
        ///// <summary>
        ///// 發行人_穆迪長期本國銀行存款評等 (發行人\ moody\國內)
        ///// </summary>
        //[SetDateFiled("RTG_MDY_LT_LC_BANK_DEP_RTG_DT")]
        //public string RTG_MDY_LOCAL_LT_BANK_DEPOSITS { get; set; }

        ///// <summary>
        ///// 發行人_穆迪長期本國銀行存款評等日期
        ///// </summary>
        //public string RTG_MDY_LT_LC_BANK_DEP_RTG_DT { get; set; }

        /// <summary>
        /// 惠譽長期發行人違約評等 (發行人\ Fitch\國外)
        /// </summary>
        [SetDateFiled("RTG_FITCH_LT_ISSUER_DFLT_RTG_DT")]
        public string RTG_FITCH_LT_ISSUER_DEFAULT { get; set; }

        /// <summary>
        /// 惠譽長期發行人違約評等日期
        /// </summary>
        public string RTG_FITCH_LT_ISSUER_DFLT_RTG_DT { get; set; }

        /// <summary>
        /// 發行人_惠譽優先無擔保債務評等 (發行人\ Fitch\國外)
        /// </summary>
        [SetDateFiled("RTG_FITCH_SEN_UNSEC_RTG_DT")]
        public string RTG_FITCH_SEN_UNSECURED { get; set; }

        /// <summary>
        /// 發行人_惠譽優先無擔保債務評等日期
        /// </summary>
        public string RTG_FITCH_SEN_UNSEC_RTG_DT { get; set; }

        /// <summary>
        /// 惠譽長期外幣發行人違約評等 (發行人\ Fitch\國外)
        /// </summary>
        [SetDateFiled("RTG_FITCH_LT_FC_ISS_DFLT_RTG_DT")]
        public string RTG_FITCH_LT_FC_ISSUER_DEFAULT { get; set; }

        /// <summary>
        /// 惠譽長期外幣發行人違約評等日期
        /// </summary>
        public string RTG_FITCH_LT_FC_ISS_DFLT_RTG_DT { get; set; }

        /// <summary>
        /// 惠譽長期本國貨幣發行人違約評等 (發行人\ Fitch\國外)
        /// </summary>
        [SetDateFiled("RTG_FITCH_LT_LC_ISS_DFLT_RTG_DT")]
        public string RTG_FITCH_LT_LC_ISSUER_DEFAULT { get; set; }

        /// <summary>
        /// 惠譽長期本國貨幣發行人違約評等日期
        /// </summary>
        public string RTG_FITCH_LT_LC_ISS_DFLT_RTG_DT { get; set; }

        /// <summary>
        /// 發行人_惠譽國內長期評等 (發行人\ Fitch(twn)\國內)
        /// </summary>
        [SetDateFiled("RTG_FITCH_NATIONAL_LT_DT")]
        public string RTG_FITCH_NATIONAL_LT { get; set; }

        /// <summary>
        /// 發行人_惠譽國內長期評等日期
        /// </summary>
        public string RTG_FITCH_NATIONAL_LT_DT { get; set; }

        /// <summary>
        /// 發行人_TRC 長期評等 (發行人\ CW\國內)
        /// </summary>
        [SetDateFiled("RTG_TRC_LONG_TERM_RTG_DT")]
        public string RTG_TRC_LONG_TERM { get; set; }

        /// <summary>
        /// 發行人_TRC 長期評等日期
        /// </summary>
        public string RTG_TRC_LONG_TERM_RTG_DT { get; set; }

        /// <summary>
        /// 標普長期外幣保證人信用評等 (保證人\ sp\國外)
        /// </summary>
        [SetDateFiled("G_RTG_SP_LT_FC_ISS_CRED_RTG_DT")]
        public string G_RTG_SP_LT_FC_ISSUER_CREDIT { get; set; }

        /// <summary>
        /// 標普長期外幣保證人信用評等日期
        /// </summary>
        public string G_RTG_SP_LT_FC_ISS_CRED_RTG_DT { get; set; }

        /// <summary>
        /// 標普本國貨幣長期保證人信用評等 (保證人\ sp\國外)
        /// </summary>
        [SetDateFiled("G_RTG_SP_LT_LC_ISS_CRED_RTG_DT")]
        public string G_RTG_SP_LT_LC_ISSUER_CREDIT { get; set; }

        /// <summary>
        /// 標普本國貨幣長期保證人信用評等日期
        /// </summary>
        public string G_RTG_SP_LT_LC_ISS_CRED_RTG_DT { get; set; }

        /// <summary>
        /// 穆迪保證人評等 (保證人\ moody\國外)
        /// </summary>
        [SetDateFiled("G_RTG_MDY_ISSUER_RTG_DT")]
        public string G_RTG_MDY_ISSUER { get; set; }

        /// <summary>
        /// 穆迪保證人評等日期
        /// </summary>
        public string G_RTG_MDY_ISSUER_RTG_DT { get; set; }

        /// <summary>
        /// 保證人_穆迪長期評等 (保證人\ moody\國外)
        /// </summary>
        [SetDateFiled("G_RTG_MOODY_LONG_TERM_DATE")]
        public string G_RTG_MOODY_LONG_TERM { get; set; }

        /// <summary>
        /// 保證人_穆迪長期評等日期
        /// </summary>
        public string G_RTG_MOODY_LONG_TERM_DATE { get; set; }

        /// <summary>
        /// 保證人_穆迪優先無擔保債務評等 (保證人\ moody\國外)
        /// </summary>
        [SetDateFiled("G_RTG_MDY_SEN_UNSEC_RTG_DT")]
        public string G_RTG_MDY_SEN_UNSECURED_DEBT { get; set; }

        /// <summary>
        /// 保證人_穆迪優先無擔保債務評等_日期
        /// </summary>
        public string G_RTG_MDY_SEN_UNSEC_RTG_DT { get; set; }

        /// <summary>
        /// 穆迪外幣保證人評等 (保證人\ moody\國外)
        /// </summary>
        [SetDateFiled("G_RTG_MDY_FC_CURR_ISSUER_RTG_DT")]
        public string G_RTG_MDY_FC_CURR_ISSUER_RATING { get; set; }

        /// <summary>
        /// 穆迪外幣保證人評等日期
        /// </summary>
        public string G_RTG_MDY_FC_CURR_ISSUER_RTG_DT { get; set; }


        //2017/12/25 A59 欄位不顯示
        ///// <summary>
        ///// 保證人_穆迪長期本國銀行存款評等 (保證人\ moody\國內)
        ///// </summary>
        //[SetDateFiled("G_RTG_MDY_LT_LC_BANK_DEP_RTG_DT")]
        //public string G_RTG_MDY_LOCAL_LT_BANK_DEPOSITS { get; set; }

        ///// <summary>
        ///// 保證人_穆迪長期本國銀行存款評等日期
        ///// </summary>
        //public string G_RTG_MDY_LT_LC_BANK_DEP_RTG_DT { get; set; }

        /// <summary>
        /// 惠譽長期保證人違約評等 (保證人\ Fitch\國外)
        /// </summary>
        [SetDateFiled("G_RTG_FITCH_LT_ISSUER_DFLT_RTG_DT")]
        public string G_RTG_FITCH_LT_ISSUER_DEFAULT { get; set; }

        /// <summary>
        /// 惠譽長期保證人違約評等日期
        /// </summary>
        public string G_RTG_FITCH_LT_ISSUER_DFLT_RTG_DT { get; set; }

        /// <summary>
        /// 保證人_惠譽優先無擔保債務評等 (保證人\ Fitch\國外)
        /// </summary>
        [SetDateFiled("G_RTG_FITCH_SEN_UNSEC_RTG_DT")]
        public string G_RTG_FITCH_SEN_UNSECURED { get; set; }

        /// <summary>
        /// 保證人_惠譽優先無擔保債務評等日期
        /// </summary>
        public string G_RTG_FITCH_SEN_UNSEC_RTG_DT { get; set; }

        /// <summary>
        /// 惠譽長期外幣保證人違約評等 (保證人\ Fitch\國外)
        /// </summary>
        [SetDateFiled("G_RTG_FITCH_LT_FC_ISS_DFLT_RTG_DT")]
        public string G_RTG_FITCH_LT_FC_ISSUER_DEFAULT { get; set; }

        /// <summary>
        /// 惠譽長期外幣保證人違約評等日期
        /// </summary>
        public string G_RTG_FITCH_LT_FC_ISS_DFLT_RTG_DT { get; set; }

        /// <summary>
        /// 惠譽長期本國貨幣保證人違約評等 (保證人\ Fitch\國外)
        /// </summary>
        [SetDateFiled("G_RTG_FITCH_LT_LC_ISS_DFLT_RTG_DT")]
        public string G_RTG_FITCH_LT_LC_ISSUER_DEFAULT { get; set; }

        /// <summary>
        /// 惠譽長期本國貨幣保證人違約評等日期
        /// </summary>
        public string G_RTG_FITCH_LT_LC_ISS_DFLT_RTG_DT { get; set; }

        /// <summary>
        /// 保證人_惠譽國內長期評等 (保證人\ Fitch(twn)\國內)
        /// </summary>
        [SetDateFiled("G_RTG_FITCH_NATIONAL_LT_DT")]
        public string G_RTG_FITCH_NATIONAL_LT { get; set; }

        /// <summary>
        /// 保證人_惠譽國內長期評等日期
        /// </summary>
        public string G_RTG_FITCH_NATIONAL_LT_DT { get; set; }

        /// <summary>
        /// 保證人_TRC 長期評等 (保證人\ CW\國內)
        /// </summary>
        [SetDateFiled("G_RTG_TRC_LONG_TERM_RTG_DT")]
        public string G_RTG_TRC_LONG_TERM { get; set; }

        /// <summary>
        /// 保證人_TRC 長期評等日期
        /// </summary>
        public string G_RTG_TRC_LONG_TERM_RTG_DT { get; set; }

        /// <summary>
        /// 債項最終評等
        /// </summary>
        [Description("債項最終評等")]
        public string Bonds_Rating { get; set; }

        /// <summary>
        /// 發行者最終評等
        /// </summary>
        [Description("發行者最終評等")]
        public string ISSUER_Rating { get; set; }

        /// <summary>
        /// 擔保者最終評等
        /// </summary>
        [Description("擔保者最終評等")]
        public string GUARANTOR_Rating { get; set; }

        /// <summary>
        /// 最終評等
        /// </summary>
        [Description("最終評等")]
        public string All_Rating { get; set; }

        /// <summary>
        /// Fill_up_YN
        /// </summary>
        [Description("Fill_up_YN")]
        public string Fill_up_YN { get; set; }
    }
}