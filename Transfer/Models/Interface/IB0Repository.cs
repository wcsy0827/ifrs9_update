using System;
using System.Collections.Generic;
using Transfer.Utility;
using Transfer.ViewModels;

namespace Transfer.Models.Interface
{
    public interface IB0Repository
    {
        List<string> getB06PRJID(string productCode);
        List<string> getB06FLOWID(string productCode, string PRJID);
        List<string> getB06CompID(string productCode, string PRJID, string FLOWID);
        Tuple<bool, List<B06ViewModel>> getB06All(B06ViewModel dataModel);
        Tuple<bool, List<B06ViewModel>> getB06(B06ViewModel dataModel);
        MSGReturnModel saveB06(string actionType, B06ViewModel dataModel);
        MSGReturnModel deleteB06(B06ViewModel dataModel);
    }
}