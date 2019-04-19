using System;
using System.Collections.Generic;
using Transfer.Utility;
using Transfer.ViewModels;

namespace Transfer.Models.Interface
{
    public interface IC2Repository
    {
        Tuple<bool, List<C13ViewModel>, List<jqGridColModel>, List<string>> getC13LagData(string tableName);
        MSGReturnModel DownLoadC13Excel(string path, List<C13ViewModel> dbDatas, List<jqGridColModel> jqgridColModel, List<string> jqgridColNames);
        Tuple<bool, List<C23ViewModel>> getC13Column();
        MSGReturnModel transferC13(string lagNumber, string tableName, List<C23ViewModel> C23);

        Tuple<bool, List<C03ViewModel>, List<jqGridColModel>, List<string>> getC03LagData(string tableName);
        MSGReturnModel DownLoadC03Excel(string path, List<C03ViewModel> dbDatas, List<jqGridColModel> jqgridColModel, List<string> jqgridColNames);
        Tuple<bool, List<C23ViewModel>> getC03Column();
        MSGReturnModel transferC03(string lagNumber, string tableName, List<C23ViewModel> C23);

        Tuple<bool, List<C04ViewModel>, List<jqGridColModel>, List<string>> getC04LagData(string tableName);
        MSGReturnModel DownLoadC04Excel(string path, List<C04ViewModel> dbDatas, List<jqGridColModel> jqgridColModel, List<string> jqgridColNames);
        Tuple<bool, List<C23ViewModel>> getC04Column();
        MSGReturnModel transferC04(string lagNumber, string tableName, List<C23ViewModel> C23);
    }
}