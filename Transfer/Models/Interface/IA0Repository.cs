using System;
using System.Collections.Generic;
using Transfer.ViewModels;

namespace Transfer.Models.Interface
{
    public interface IA0Repository
    {
        Tuple<bool, List<A08ViewModel>> queryA08(string referenceNbr, string reportDate);
    }
}