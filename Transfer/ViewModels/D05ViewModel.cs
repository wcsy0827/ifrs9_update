using System;
using System.Runtime.Serialization;

namespace Transfer.ViewModels
{
    /// <summary>
    /// D05 View Data
    /// </summary>
    [DataContract]
    [Serializable]
    public class D05ViewModel
    {
        [DataMember]
        public string Group_Product { get; set; }

        [DataMember]
        public string Group_Product_Code { get; set; }

        [DataMember]
        public string Processing_Date { get; set; }

        [DataMember]
        public string Product_Code { get; set; }
    }
}