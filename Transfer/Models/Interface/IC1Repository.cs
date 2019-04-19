using System;
using System.Collections.Generic;
using Transfer.Utility;
using Transfer.ViewModels;

namespace Transfer.Models.Interface
{
    public interface IC1Repository
    {
        Tuple<bool, List<C13ViewModel>> getC13All();
        Tuple<bool, List<C13ViewModel>> getC13(C13ViewModel dataModel);
        MSGReturnModel saveC13(string actionType, C13ViewModel dataModel);
        MSGReturnModel deleteC13(string yearQuartly);
    }
}