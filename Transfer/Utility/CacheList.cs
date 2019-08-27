namespace Transfer.Utility
{
    /// <summary>
    /// cache命名 目的為不重複cache名稱 避免資料被覆蓋
    /// </summary>
    public static class CacheList
    {
        #region 資料庫資料
        public static string A08DbfileData { get; private set; }
        public static string A41DbfileData { get; private set; }
        public static string A41ExcelfileData { get; private set; }
        public static string A42DbfileData { get; private set; }
        public static string A42ExcelfileData { get; private set; }
        public static string A44DbfileData { get; private set; }
        public static string A46DbfileData { get; private set; }
        public static string A46ExcelfileData { get; private set; }
        public static string A47DbfileData { get; private set; }
        public static string A47ExcelfileData { get; private set; }
        public static string A48DbfileData { get; private set; }
        public static string A48ExcelfileData { get; private set; }
        public static string A49DbfileData { get; private set; }
        public static string A49ExcelfileData { get; private set; }
        public static string A52DbfileData { get; private set; }
        public static string A52AuditDbfileData { get; private set; }
        public static string A52AuditDetailDbfileData { get; private set; }
        public static string A56DbfileData { get; private set; }
        public static string A57DbfileData { get; private set; }
        public static string A58DbfileData { get; private set; }
        public static string A58DbMissfileData { get; private set; }
        public static string A59ExcelfileData { get; private set; }
        public static string A94DbfileData { get; private set; }
        public static string A95DbfileData { get; private set; }
        public static string A95ExcelfileData { get; private set; }
        public static string A95_1DbfileData { get; private set; }
        public static string A95_1ExcelfileData { get; private set; }
        public static string A96DbfileData { get; private set; }
        public static string A96ExcelfileData { get;private set; }
        public static string A96TradeDbfileData { get; private set; }
        public static string B06DbfileData { get; private set; }
        public static string C03DbfileData { get; private set; }
        public static string C04DbfileData { get; private set; }
        public static string C04_1DbfileData { get; private set; }
        public static string C04TransferData { get; private set; }
        public static string C10DbfileData { get; private set; }
        public static string C10ExcelfileData { get; set; }
        public static string C13DbfileData { get; private set; }
        public static string D53DbfileData { get; private set; }
        public static string D53ExcelfileData { get; private set; }
        public static string D54InsertSearchData { get;private set; }
        public static string D54DbfileData { get; private set; }
        public static string D03DbfileData { get; private set; }
        public static string D60DbfileData { get; private set; }
        public static string D61DbfileData { get; private set; }
        public static string D63SearchCache { get; private set; }
        public static string D63DbfileTransferData { get; private set; }
        public static string D63DbfileData { get; private set; }
        public static string D63DbfileHistoryData { get; private set; }
        public static string D64DbfileData { get; private set; }
        public static string D65SearchCache { get; private set; }
        public static string D65DbfileTransferData { get; private set; }
        public static string D65DbfileData { get; private set; }
        public static string D65DbfileHistoryData { get; private set; }
        public static string D66DbfileData { get; private set; }
        public static string D66DbfileDetailData { get; private set; }
        public static string D68DbfileData { get; private set; }
        public static string D72DbfileData { get; private set; }
        public static string D72ExcelfileData { get; private set; }
        public static string D73DbfileData { get; private set; }
        public static string D74DbfileData { get; private set;}
        public static string D74_1DbfileData { get; private set; }
       
        public static string CheckTableDbfileData { get; private set; }
        public static string CheckTableDbfileData2 { get; private set; }
        public static string CheckData { get; private set; }
        public static string userDbData { get; set; }
        public static string userLogInUserDbData { get; set; }
        public static string userLogInBrowserDbData { get; set; }
        public static string userLogInEventDbData { get; set; }
        public static string SetAssessment { get; set; }
        public static string SetAssessmentAuditor { get; set; }
        public static string SetAssessmentPresented { get; set; }
        public static string D62DbfileData { get; private set; }
        public static string C01DbfileData { get; private set; }
        public static string C01ExcelfileData { get; private set; }
        public static string C07DbfileData { get; private set; }
        public static string C07AdvancedDbfileData { get; private set; }
        public static string C07AdvancedSumDbfileData { get; private set; }
        public static string D67DbfileData { get; private set; }
        public static string D02DbfileData { get; private set; }
        public static string MenuMainDbfileData { get; set; }
        public static string MenuSubDbfileData { get; set; }
        public static string ReELDbfileData { get; set; }


        #endregion 資料庫資料

        #region 上傳檔名

        public static string A41ExcelName { get; private set; }

        public static string A42ExcelName { get; private set; }

        public static string A45ExcelName { get; private set; }

        public static string A46ExcelName { get; private set; }

        public static string A47ExcelName { get; private set; }

        public static string A48ExcelName { get; private set; }

        public static string A49ExcelName { get; private set; }

        public static string A59ExcelName { get; private set; }

        public static string A62ExcelName { get; private set; }

        public static string A63ExcelName { get; private set; }

        public static string A71ExcelName { get; private set; }

        public static string A81ExcelName { get; private set; }

        public static string A95ExcelName { get; private set; }

        public static string A96ExcelName { get; private set; }
        public static string A95_1ExcelName { get; private set; }
        public static string C01ExcelName { get; private set; }
        public static string D53ExcelName { get; private set; }
        public static string D72ExcelName { get; private set; }
        public static string C10ExcelName { get; private set; }

        #endregion 上傳檔名

        static CacheList()
        {
            #region 資料庫資料

            A08DbfileData = "A08DbfileData";
            A41DbfileData = "A41DbfileData";
            A41ExcelfileData = "A41ExcelfileData";
            A42DbfileData = "A42DbfileData";
            A42ExcelfileData = "A42ExcelfileData";
            A44DbfileData = "A44DbfileData";
            A46DbfileData = "A46DbfileData";
            A46ExcelfileData = "A46ExcelfileData";
            A47DbfileData = "A47DbfileData";
            A47ExcelfileData = "A47ExcelfileData";
            A48DbfileData = "A48DbfileData";
            A48ExcelfileData = "A48ExcelfileData";
            A49DbfileData = "A49DbfileData";
            A49ExcelfileData = "A49ExcelfileData";
            A52DbfileData = "A52DbfileData";
            A52AuditDbfileData = "A52AuditDbfileData";
            A52AuditDetailDbfileData = "A52AuditDetailDbfileData";
            A56DbfileData = "A56DbfileData";
            A57DbfileData = "A57DbfileData";
            A58DbfileData = "A58DbfileData";
            A58DbMissfileData = "A58DbMissfileData";
            A59ExcelfileData = "A59ExcelfileData";
            A94DbfileData = "A94DbfileData";
            A95DbfileData = "A95DbfileData";
            A95ExcelfileData = "A95ExcelfileData";
            A95_1DbfileData = "A95_1DbfileData";
            A95_1ExcelfileData = "A95_1ExcelfileData";
            A96DbfileData = "A96DbfileData";
            A96ExcelfileData = "A96ExcelfileData";
            A96TradeDbfileData = "A96TradeDbfileData";
            B06DbfileData = "B06DbfileData";
            C03DbfileData = "C03DbfileData";
            C04DbfileData = "C04DbfileData";
            C04_1DbfileData = "C04_1DbfileData";
            C04TransferData = "C04TransferData";
            C10DbfileData = "C10DbfileData";
            C13DbfileData = "C13DbfileData";
            D03DbfileData = "D03DbfileData";
            D53DbfileData = "D53DbfileData";
            D53ExcelfileData = "D53ExcelfileData";
            D54InsertSearchData = "D54InsertSearchData";
            D54DbfileData = "D54DbfileData";
            D60DbfileData = "D60DbfileData";
            D61DbfileData = "D61DbfileData";
            D63SearchCache = "D63SearchCache";
            D63DbfileTransferData = "D63DbfileTransferData";
            D63DbfileData = "D63DbfileData";
            D63DbfileHistoryData = "D63DbfileHistoryData";
            D64DbfileData = "D64DbfileData";
            D65SearchCache = "D65SearchCache";
            D65DbfileTransferData = "D65DbfileTransferData";
            D65DbfileData = "D65DbfileData";
            D65DbfileHistoryData = "D65DbfileHistoryData";
            D66DbfileData = "D66DbfileData";
            D66DbfileDetailData = "D66DbfileDetailData";
            D68DbfileData = "D68DbfileData";
            D72DbfileData = "D72DbfileData";
            D72ExcelfileData = "D72ExcelfileData";
            D73DbfileData = "D73DbfileData";
            D74DbfileData = "D74DbfileData";
            D74_1DbfileData = "D74_1DbfileData";
            C10ExcelfileData = "E01ExcelfileData";
            CheckTableDbfileData = "CheckTableDbfileData";
            CheckTableDbfileData2 = "CheckTableDbfileData2";
            CheckData = "CheckData";
            userDbData = "userDbData";
            userLogInUserDbData = "userLogInUserDbData";
            userLogInBrowserDbData = "userLogInBrowserDbData";
            userLogInEventDbData = "userLogInEventDbData";
            SetAssessment = "SetAssessment";
            SetAssessmentAuditor = "SetAssessmentAuditor";
            SetAssessmentPresented = "SetAsse;ssmentPresented";
            D62DbfileData = "D62DbfileData";
            C01ExcelfileData = "C01ExcelfileData";
            C01DbfileData = "C01DbfileData";
            C07DbfileData = "C07DbfileData";
            C07AdvancedDbfileData = "C07AdvancedDbfileData";
            C07AdvancedSumDbfileData = "C07AdvancedSumDbfileData";
            D67DbfileData = "D67DbfileData";
            D02DbfileData = "D02DbfileData";
            MenuMainDbfileData = "MenuMainDbfileData";
            MenuSubDbfileData = "MenuSubDbfileData";
            ReELDbfileData = "ReELDbfileData";
            

            #endregion 資料庫資料

            #region 上傳檔名

            A41ExcelName = "A41ExcelName";
            A42ExcelName = "A42ExcelName";
            A45ExcelName = "A45ExcelName";
            A46ExcelName = "A46ExcelName";
            A47ExcelName = "A47ExcelName";
            A48ExcelName = "A48ExcelName";
            A49ExcelName = "A49ExcelName";
            A59ExcelName = "A59ExcelName";
            A62ExcelName = "A62ExcelName";
            A63ExcelName = "A63ExcelName";
            A71ExcelName = "A71ExcelName";
            A81ExcelName = "A81ExcelName";
            A95ExcelName = "A95ExcelName";
            A95_1ExcelName = "A95_1ExcelName";
            A96ExcelName = "A96ExcelName";
            C01ExcelName = "C01ExcelName";
            D53ExcelName = "D53ExcelName";
            D72ExcelName = "D72ExcelName";
            C10ExcelName = "C10ExcelName";
            #endregion 上傳檔名
        }
    }
}