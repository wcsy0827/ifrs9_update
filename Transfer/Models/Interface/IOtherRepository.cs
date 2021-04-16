using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using Transfer.Utility;
using Transfer.ViewModels;

namespace Transfer.Models.Interface
{
    public interface IOtherRepository
    {
        List<CheckTableViewModel> getRatingSchedule(DateTime? dt);
        List<SelectOption> getReportVersion(DateTime reportdate);
    }
}