using System.Collections.Generic;

namespace Transfer.Utility
{
    public class GetQueryValue
    {
    }

    public class VersionInfo
    {
        public string Version { get; set; }
        public string VersionContent { get; set; }
        public string VersionData()
        {
            return Version + "_" + VersionContent;
        }
    }

    public class StatusInfo
    {
        public string Status { get; set; }
        public string Description { get; set; }
        public string StatusData()
        {
            return Status + "_" + Description;
        }
    }

    public class GroupProduct
    {
        public string GroupProductName { get; set; }
        public string GroupProductCode { get; set; }
        public IEnumerable<string> ProductCode { get; set; }
        public string ProductData()
        {
            string productCode = string.Join(",", ProductCode);
            return productCode + "|" + GroupProductName + " (" + GroupProductCode + ")";
        }
    }

    public class BondBasicAssessment
    {
        public string RuleID { get; set; }
        public string RuleDesc { get; set; }
        public string NumberPen { get; set; }
        public string BondBasicAssessmentData()
        {
            if (RuleDesc == null) { RuleDesc = "無說明"; }
            return "規則編號:" + RuleID + "(" + RuleDesc + "):" + NumberPen + "筆";
        }
    }

    public class BondRiskReviewResult
    {
        public string Reference_Nbr { get; set; }
        public string Create_User { get; set; }
        public string First_Order_User { get; set; }
        public string Second_Order_User { get; set; }
        public string First_Order_Confirm { get; set; }
        public string Second_Order_Confirm { get; set; }
        public string D62Handle { get; set; }
        public string D62HandleOpinion { get; set; }
        public string D63_D64Handle { get; set; }
        public string D63_D64HandleOpinion { get; set; }
        public string SummaryHandle { get; set; }
        public string SummaryHandleOpinion { get; set; }
        public string WatchINDHandle { get; set; }
        public string WatchINDHandleOpinion { get; set; }
        public string WarningINDHandle { get; set; }
        public string WarningINDHandleOpinion { get; set; }
        public string C07AdvancedHandle { get; set; }
        public string C07AdvancedHandleOpinion { get; set; }
        public string D62ReviewOne { get; set; }
        public string D62ReviewOneOpinion { get; set; }
        public string D63_D64ReviewOne { get; set; }
        public string D63_D64ReviewOneOpinion { get; set; }
        public string SummaryReviewOne { get; set; }
        public string SummaryReviewOneOpinion { get; set; }
        public string WatchINDReviewOne { get; set; }
        public string WatchINDReviewOneOpinion { get; set; }
        public string WarningINDReviewOne { get; set; }
        public string WarningINDReviewOneOpinion { get; set; }
        public string C07AdvancedReviewOne { get; set; }
        public string C07AdvancedReviewOneOpinion { get; set; }
        public string D62ReviewTwo { get; set; }
        public string D62ReviewTwoOpinion { get; set; }
        public string D63_D64ReviewTwo { get; set; }
        public string D63_D64ReviewTwoOpinion { get; set; }
        public string SummaryReviewTwo { get; set; }
        public string SummaryReviewTwoOpinion { get; set; }
        public string WatchINDReviewTwo { get; set; }
        public string WatchINDReviewTwoOpinion { get; set; }
        public string WarningINDReviewTwo { get; set; }
        public string WarningINDReviewTwoOpinion { get; set; }
        public string C07AdvancedReviewTwo { get; set; }
        public string C07AdvancedReviewTwoOpinion { get; set; }
    }
}