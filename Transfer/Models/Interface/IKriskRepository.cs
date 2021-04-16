using System;
using System.Collections.Generic;
using Transfer.Utility;
using Transfer.ViewModels;
using static Transfer.Enum.Ref;

namespace Transfer.Models.Interface
{
    public interface IKriskRepository
    {
        List<SelectOption> getProduct(GroupProductCode type);

        Tuple<string, string> getProductInfo(string productCode);

        List<string> GetC01Version(string product, DateTime date);

        void saveD04(Flow_Apply_Status data);

        Tuple<string,List<EL_Data_Out>> deleteC07(Flow_Apply_Status data);

        bool checkC07(DateTime reportDate, string PRJID, string FLOWID);

        void backC07(List<EL_Data_Out> datas);

        void getBondC07CheckData(DateTime dt, int version);

        //(190628 John.換券應收未收金額修正)
        //執行UniRisk前檢核該報導日版本是否已經進入風控覆核流程
        //(190628 John.換券應收未收金額修正)
        //執行UniRisk前檢核該報導日版本是否已經進入風控覆核流程
        MSGReturnModel CheckInfo(string date, string version);
    }
}