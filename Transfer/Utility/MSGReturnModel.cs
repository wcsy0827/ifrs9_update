﻿using System.ComponentModel;
using System.Runtime.Serialization;
using System.Web.Mvc;

namespace Transfer.Utility
{
    /// <summary>
    /// 傳到前端的Model
    /// </summary>
    [DataContract]
    public class MSGReturnModel
    {
        [DataMember]
        public JsonResult colModel { get; set; }

        [DataMember]
        public JsonResult colNames { get; set; }

        /// <summary>
        /// 回傳資料
        /// </summary>
        [DataMember]
        public JsonResult Datas { get; set; }

        /// <summary>
        /// 回傳訊息
        /// </summary>
        [DataMember]
        [DisplayName("Message Description")]
        public string DESCRIPTION { get; set; }

        /// <summary>
        /// ReasonCode
        /// </summary>
        [DataMember]
        [DisplayName("Message Reason Code")]
        public string REASON_CODE { get; set; }

        /// <summary>
        /// 是否成功
        /// </summary>
        [DataMember]
        [DisplayName("Message Return Flag")]
        public bool RETURN_FLAG { get; set; }
    }
}