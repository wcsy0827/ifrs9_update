using System;
using System.Collections.Generic;
using System.Linq;
using Transfer.Models.Interface;
using Transfer.Utility;
using Transfer.ViewModels;

namespace Transfer.Models.Repository
{
    public class A0Repository : IA0Repository
    {
        #region getA08
        public Tuple<bool, List<A08ViewModel>> queryA08(string referenceNbr, string reportDate)
        {
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                if (db.Loan_Report_Info.Any())
                {
                    var query = from q in db.Loan_Report_Info.AsNoTracking()
                                select q;

                    if (referenceNbr.IsNullOrWhiteSpace() == false)
                    {
                        query = query.Where(x => x.Reference_Nbr == referenceNbr);
                    }

                    if (reportDate.IsNullOrWhiteSpace() == false)
                    {
                        DateTime tempDate = DateTime.Parse(reportDate);
                        query = query.Where(x => x.Report_Date == tempDate);
                    }

                    return new Tuple<bool, List<A08ViewModel>>(query.Any(),
                                                               query.AsEnumerable()
                                                               .Select(x => { return DbToA08ViewModel(x); }).ToList());
                }
            }

            return new Tuple<bool, List<A08ViewModel>>(false, new List<A08ViewModel>());
        }

        #endregion

        #region Db 組成 A08ViewModel
        private A08ViewModel DbToA08ViewModel(Loan_Report_Info data)
        {
            return new A08ViewModel()
            {
                Reference_Nbr = data.Reference_Nbr,
                Processing_Date = data.Processing_Date.ToString("yyyy/MM/dd"),
                Report_Date = data.Report_Date.ToString("yyyy/MM/dd"),
                Loan_Risk_Type = data.Loan_Risk_Type,
                Interest = data.Interest.ToString()
            };
        }
        #endregion
    }
}