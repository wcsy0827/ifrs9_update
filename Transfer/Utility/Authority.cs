using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using Transfer.Models;
using static Transfer.Enum.Ref;

namespace Transfer.Utility
{
    public class Authority
    {
        public Authority(string tableId ,string productCode = null)
        {
            if (productCode.IsNullOrWhiteSpace())
                productCode = Assessment_Type.B.GetDescription();
            var _common = new Controllers.CommonController();
            userAccount = Controllers.AccountController.CurrentUserInfo.Name;
            Presenteds = _common.getAssessment(productCode, tableId, SetAssessmentType.Presented);
            Auditors = _common.getAssessment(productCode, tableId, SetAssessmentType.Auditor);
            if (Presenteds.Any(x => x.User_Account == userAccount))
            {
                AuthorityType = Authority_Type.Presented;
            }
            else if (Auditors.Any(x => x.User_Account == userAccount))
            {
                AuthorityType = Authority_Type.Auditor;
            }
            else
            {
                AuthorityType = Authority_Type.None;
            }
        }
        /// <summary>
        /// 可呈送複核者
        /// </summary>
        public List<IFRS9_User> Presenteds { get; private set; }

        /// <summary>
        /// 複核者
        /// </summary>
        public List<IFRS9_User> Auditors { get; private set; }

        /// <summary>
        /// 
        /// </summary>
        public Authority_Type AuthorityType { get;private set; }

        public string userAccount { get; private set; }
    }
}