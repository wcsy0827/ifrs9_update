﻿using System.ComponentModel;

namespace Transfer.ViewModels
{
    public class BondRatingViewModel
    {
        [Description("資料表名稱")]
        public string Table_Name { get; set; }

        [Description("編號")]
        public string Bond_Rating_ID { get; set; }

        [Description("Bond_Number")]
        public string Bond_Number { get; set; }

        [Description("S&P")]
        public string S_And_P { get; set; }

        [Description("Moody's")]
        public string Moodys { get; set; }

        [Description("Fitch")]
        public string Fitch { get; set; }

        [Description("Fitch_TW")]
        public string Fitch_TW { get; set; }

        [Description("TRC")]
        public string TRC { get; set; }
    }
}