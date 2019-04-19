using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using Transfer.ViewModels;

namespace Transfer.Models.Interface
{
    public interface IOtherRepository
    {
        List<CheckTableViewModel> getRatingSchedule(DateTime? dt);
    }
}