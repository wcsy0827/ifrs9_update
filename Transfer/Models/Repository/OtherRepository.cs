using System;
using System.Collections.Generic;
using System.Linq;
using Transfer.Models.Interface;
using Transfer.Utility;
using Transfer.ViewModels;
using static Transfer.Enum.Ref;

namespace Transfer.Models.Repository
{
    public class OtherRepository : IOtherRepository
    {
        public List<CheckTableViewModel> getRatingSchedule(DateTime? dt)
        {
            List<string> fileNames = new List<string>() {
                Table_Type.A53.ToString(),
                Table_Type.A57.ToString(),
                Table_Type.A58.ToString(),
                Table_Type.A07.ToString(),
                Table_Type.A84.ToString(),
                Table_Type.A91.ToString(),
                Table_Type.A92.ToString(),
                Table_Type.A93.ToString(),
                Table_Type.A961.ToString(),
                Table_Type.A962.ToString(),
                Table_Type.C03.ToString(),
                Table_Type.C04.ToString()
            };

            List<CheckTableViewModel> data = new List<CheckTableViewModel>();
            try
            {
                using (IFRS9DBEntities db = new IFRS9DBEntities())
                {
                    var tc = db.Transfer_CheckTable.AsNoTracking().Where(x =>
                    x.Create_date != null &&
                    x.Create_time != null &&
                    x.Process == "Plan" && //只要顯示排程的
                    fileNames.Contains(x.File_Name));

                    foreach (var item in tc.GroupBy(z => z.File_Name))
                    {
                        var x = item
                            .OrderByDescending(y => y.Create_date)
                            .ThenByDescending(y => y.Create_time)
                            .First();
                        data.Add(new CheckTableViewModel()
                        {
                            ReportDate = x.ReportDate.ToString("yyyy/MM/dd"),
                            Version = x.Version.ToString(),
                            TransferType = x.TransferType,
                            Create_Date = x.Create_date,
                            Create_Time = x.Create_time,
                            End_Date = x.End_date,
                            End_Time = x.End_time,
                            File_Name = $"{x.File_Name} ({x.File_Name.tableNameGetDescription()})"
                        });
                    }
                }
            }
            catch
            {
            }
            return data;
        }


        public List<SelectOption> getReportVersion(DateTime reportdate)
        {
            List<SelectOption> result = new List<SelectOption>();
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                var vers = db.Version_Info.Where(x => x.Report_Date == reportdate&&x.Risk_Control_Status==5).Select(x => x.Version).DefaultIfEmpty(0).Distinct().Max();
                if (vers!=0)
                {

                    var selectoption = new SelectOption() { Value = vers.ToString(), Text = vers.ToString() };
                    result.Insert(result.Count(), selectoption);

                }
            }
            return result;
        }
    }
}