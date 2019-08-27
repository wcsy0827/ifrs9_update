using System;
using System.Collections.Generic;
using System.ComponentModel;
using Transfer.Enum;
using Transfer.Models;
using static Transfer.Enum.Ref;

namespace Transfer.ViewModels
{
    public class C10ViewModel_1 //由C10ViewModel轉換而來
    {

        public Bond_Account_AssessmentCheck Data { get; set; }

        public bool Amort { get; set; }

        public bool Interest { get; set; }

        public bool CheckDataRedundancy(C10DataType c10Type)
        {
            bool result = false;

            switch (c10Type) {
                case C10DataType.Amort:
                    if (Amort == true)
                        result = true;
                    break;

                case C10DataType.Interest:
                    if(Interest == true)
                        result = true;
                    break;

                case C10DataType.Amort_Interest:
                    if (Interest == true || Amort == true)
                        result = true;
                    break;
                
                default:
                    break;
            }

            return result;

        }
    }
}