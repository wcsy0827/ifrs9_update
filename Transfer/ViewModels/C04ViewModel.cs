using System.ComponentModel;

namespace Transfer.ViewModels
{
    public class C04ViewModel
    {
        [Description("資料處理日期")]
        public string Processing_Date { get; set; }

        [Description("產品")]
        public string Product_Code { get; set; }

        [Description("執行資料批號")]
        public string Data_ID { get; set; }

        [Description("資料(季)")]
        public string Year_Quartly { get; set; }

        [Description("季違約資料")]
        public string PD_Quartly { get; set; }

        [Description("消費者物價指數_美國")]
        public string CPI_INDX_Index { get; set; }

        [Description("消費者物價指數_加拿大")]
        public string CACPSA_Index { get; set; }

        [Description("消費者物價指數_韓國")]
        public string KOCPI_Index { get; set; }

        [Description("消費者物價指數_歐元區")]
        public string EACPI_Index { get; set; }

        [Description("消費者物價指數_德國")]
        public string GRCP2000_Index { get; set; }

        [Description("消費者物價指數_英國")]
        public string UKRPCHVJ_Index { get; set; }

        [Description("消費者物價指數_荷蘭")]
        public string NECPIND_Index { get; set; }

        [Description("消費者物價指數_澳洲")]
        public string AUCCTOTS_Index { get; set; }

        [Description("消費者物價指數_日本")]
        public string JCPTSGEN_Index { get; set; }

        [Description("國債收益率(三個月)_美國")]
        public string USGG3M_Index { get; set; }

        [Description("國債收益率(三個月)_歐洲")]
        public string GECU3M_Index { get; set; }

        [Description("國債收益率(十年)_美國")]
        public string USGG10YR_Index { get; set; }

        [Description("國債收益率(十年)_歐洲")]
        public string GECU10YR_Index { get; set; }

        [Description("美國公共債務餘額(NSA)_美國")]
        public string DEBPTOTL_Index { get; set; }

        [Description("股票價格指數_日本")]
        public string NKY_Index { get; set; }

        [Description("股票價格指數_美國(道瓊斯工業平均指數)")]
        public string INDU_Index { get; set; }

        [Description("股票價格指數_美國(納斯達克綜合指數)")]
        public string CCMP_Index { get; set; }

        [Description("股票價格指數_美國(標準普爾500指數)")]
        public string SPX_Index { get; set; }

        [Description("股票價格指數_德國")]
        public string DAX_Index { get; set; }

        [Description("股票價格指數_歐洲")]
        public string DJST_Index { get; set; }

        [Description("失業率_美國")]
        public string USURTOT_Index { get; set; }

        [Description("失業率_歐元區")]
        public string UMRTEMU_Index { get; set; }

        [Description("失業率_日本")]
        public string JNUE_Index { get; set; }

        [Description("失業率_加拿大")]
        public string CANLXEMR_Index { get; set; }

        [Description("失業率_德國")]
        public string GRUEPR_Index { get; set; }

        [Description("失業率_英國")]
        public string UKUEILOR_Index { get; set; }

        [Description("失業率_荷蘭")]
        public string NEUENTTR_Index { get; set; }

        [Description("失業率_南韓")]
        public string KOEAUERS_Index { get; set; }

        [Description("失業率_澳洲")]
        public string AULFUNEM_Index { get; set; }

        [Description("美國聯邦住宅金融管理局-房價指數-房屋買進價格(SA)")]
        public string HPIMLEVL_Index { get; set; }

        [Description("GDP_美國")]
        public string GDP_CHWG_Index { get; set; }

        [Description("GDP_加拿大")]
        public string CGE9MP_Index { get; set; }

        [Description("GDP_歐元區")]
        public string EUGNEMU_Index { get; set; }

        [Description("GDP_德國")]
        public string GRGDEGDP_Index { get; set; }

        [Description("GDP_英國")]
        public string UKGRABMI_Index { get; set; }

        [Description("GDP_荷蘭")]
        public string NEGDPESA_Index { get; set; }

        [Description("GDP_日本")]
        public string JGDPGDP_Index { get; set; }

        [Description("GDP_南韓")]
        public string KOGDSTOT_Index { get; set; }

        [Description("GDP_澳洲")]
        public string AUNAGDP_Index { get; set; }

        [Description("Eurozone Government Debt as % of GDP(NSA)政府債務(佔GDP%)")]
        public string EUQDGEUR_Index { get; set; }

        [Description("美國經常帳")]
        public string USCABAL_Index { get; set; }

        [Description("US Current Account Balance(%GDP)")]
        public string EHCAUS_Index { get; set; }

        [Description("德國經常帳歐元(歐元，未經季調)")]
        public string GRCAEU_Index { get; set; }

        [Description("Germany Current Account Balance (% GDP)")]
        public string EHCADE_Index { get; set; }

        [Description("歐洲央行歐元區經常帳淨額 經季調")]
        public string EUSATOTN_Index { get; set; }

        [Description("EURIBOR(三個月)")]
        public string EUR003M_Index { get; set; }

        [Description("歐洲央行隔夜存款利率")]
        public string EUORDEPO_Index { get; set; }

        [Description("美國隔夜存款利率")]
        public string USDR1T_CMPN_Curncy { get; set; }

        [Description("美國聯邦基金目標利率(NSA)")]
        public string FDTR_Index { get; set; }

        [Description("ESTX石油天然氣歐元價格(Euro STOXX Oil and Gas)")]
        public string SXEE_Index { get; set; }

        [Description("NYMEX天然氣")]
        public string NG1_Comdty { get; set; }

        [Description("布蘭特原油")]
        public string CO1_COMDTY { get; set; }

        [Description("紐約商業交易所WTI原油")]
        public string CL1_Comdty { get; set; }

        [Description("紐約證交所Arca石油類指數")]
        public string XOI_Index { get; set; }

        [Description("匯率_歐洲(EURUSD_Curncy)")]
        public string EURUSD_Curncy { get; set; }

        [Description("匯率_美國(美元對人民幣間的匯率)")]
        public string USDCNY_Curncy { get; set; }

        [Description("匯率_美國(美元對日幣間的匯率)")]
        public string USDJPY_Curncy { get; set; }

        [Description("匯率_美國(美元對韓元間的匯率)")]
        public string USDKRW_Curncy { get; set; }

        [Description("匯率_美國(美元對新台幣間的匯率)")]
        public string USDTWD_Curncy { get; set; }

        [Description("匯率_歐洲(EURTWD_Curncy)")]
        public string EURTWD_Curncy { get; set; }

        [Description("匯率_中國")]
        public string CNYTWD_Curncy { get; set; }

        [Description("匯率_日本")]
        public string JPYTWD_Curncy { get; set; }

        [Description("匯率_韓國")]
        public string KRWTWD_Curncy { get; set; }

        [Description("再融資利率")]
        public string EURR002W_Index { get; set; }

        [Description("製造業採購經理人指數 經季調")]
        public string NAPMPMI_Index { get; set; }

        [Description("供應管理協會-非製造業指數(NSA)")]
        public string NAPMNMI_Index { get; set; }

        [Description("美國供應管理協會全國製造業新訂單指數 經季調(SA)")]
        public string NAPMNEWO_Index { get; set; }

        [Description("費城聯邦儲備銀行製造業企業景氣調查 - 當前總體狀況")]
        public string OUTFGAF_Index { get; set; }

        [Description("美國經濟諮詢委員會領先指標週平均首次申請失業救濟金(SA)")]
        public string LEI_WKIJ_Index { get; set; }

        [Description("Conference Board US Leading Index Leading Credit Index(SA)")]
        public string LEI_LCI_Index { get; set; }

        [Description("美國NAHB住宅市場指數(SA)")]
        public string USHBMIDX_Index { get; set; }

        [Description("Monetary Base Total NSA(貨幣基數)")]
        public string ARDIMONY_Index { get; set; }

        [Description("貨幣供給M1(美國)")]
        public string M1_Index { get; set; }

        [Description("貨幣供給M1(歐洲)")]
        public string ECMAM1_Index { get; set; }

        [Description("貨幣供給M2(美國)")]
        public string M2_Index { get; set; }

        [Description("貨幣供給M2(歐洲)")]
        public string ECMAM2_Index { get; set; }

        [Description("歐元區主要貨幣指數")]
        public string CEEREU_Index { get; set; }

        [Description("美元指數")]
        public string DXY_Curncy { get; set; }

        [Description("歐元區經濟景氣指標")]
        public string EUBCI_Index { get; set; }

        [Description("OECD德國領先經濟指標")]
        public string OEDEA044_Index { get; set; }

        [Description("歐元區經濟信心指數")]
        public string EUESEMU_Index { get; set; }

        [Description("Sentix經濟指數-歐元區總計-歐元區整體指數")]
        public string SNTEEUGX_Index { get; set; }

        [Description("歐盟統計局歐元區對外貿易貿易收支(經季調)")]
        public string XTSBEZ_Index { get; set; }

        [Description("德國貿易餘額(未經季調)(歐元)")]
        public string GRTBALE_Index { get; set; }

        [Description("ECB Survey Change in Credit Standards Lending to Business Last 3mth Net Percent(企業信貸標準)")]
        public string EBLS11NC_Index { get; set; }

        [Description("MSCI世界原物料指數")]
        public string MXWO0MT_Index { get; set; }

        [Description("標普500原物料指數")]
        public string S5MATRX_Index { get; set; }

        [Description("現貨黃金")]
        public string XAU_Curncy { get; set; }

        [Description("國際收支平衡總額")]
        public string USTBTOT_Index { get; set; }

        [Description("ADP非農人數變動(SA)")]
        public string ADP_CHNG_Index { get; set; }

        [Description("非農就業人數月比(SA)")]
        public string NFP_TCH_Index { get; set; }

        [Description("營建支出月比(SA)")]
        public string CNSTTMOM_Index { get; set; }

        [Description("工廠訂單")]
        public string TMNOCHNG_Index { get; set; }

        [Description("耐久財訂單")]
        public string DGNOCHNG_Index { get; set; }

        [Description("工業生產月比")]
        public string IP_CHNG_Index { get; set; }

        [Description("新屋開工數")]
        public string NHSPSTOT_Index { get; set; }

        [Description("新屋開工月比")]
        public string NHCHSTCH_Index { get; set; }

        [Description("營建許可數")]
        public string NHSPATOT_Index { get; set; }

        [Description("營建許可月比")]
        public string NHCHATCH_Index { get; set; }

        [Description("領先指數")]
        public string LEI_CHNG_Index { get; set; }

        [Description("成屋銷售")]
        public string ETSLTOTL_Index { get; set; }

        [Description("成屋銷售月比")]
        public string ETSLMOM_Index { get; set; }

        [Description("新屋銷售")]
        public string NHSLTOT_Index { get; set; }

        [Description("新屋銷售月比")]
        public string NHSLCHNG_Index { get; set; }

        [Description("密西根大學消費者信心指數")]
        public string CONSSENT_Index { get; set; }

        [Description("零售銷售指數")]
        public string RSTAMOM_Index { get; set; }

        [Description("生產者物價指數")]
        public string EUPPEMUM_Index { get; set; }

        [Description("消費者信心指數")]
        public string EUCCEMU_Index { get; set; }

        [Description("製造業PMI")]
        public string CPMINDX_Index { get; set; }

        [Description("Markit製造業PMI")]
        public string MPMICNMA_Index { get; set; }

        [Description("貿易收支")]
        public string CNFRBAL_Index { get; set; }

        [Description("CPI年比")]
        public string CNCPIYOY_Index { get; set; }

        [Description("PPI年比")]
        public string CHEFTYOY_Index { get; set; }

        [Description("貨幣供給M2")]
        public string CNMS2YOY_Index { get; set; }

        [Description("進口")]
        public string CNFRIMPY_Index { get; set; }

        [Description("出口")]
        public string CNFREXPY_Index { get; set; }

        [Description("德國公共財政-負債佔GDP比重(%)")]
        public string GRFIDEBT_Index { get; set; }
    }
}