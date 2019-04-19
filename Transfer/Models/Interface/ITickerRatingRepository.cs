using System;
using System.Collections.Generic;
using Transfer.Utility;
using Transfer.ViewModels;

namespace Transfer.Models.Interface
{
    public interface ITickerRatingRepository
    {
        Tuple<bool, List<GuarantorTickerViewModel>> getGuarantorTicker(string queryType, GuarantorTickerViewModel dataModel);

        MSGReturnModel saveGuarantorTicker(string actionType, GuarantorTickerViewModel dataModel);

        MSGReturnModel deleteGuarantorTicker(string issuer);

        Tuple<bool, List<IssuerTickerViewModel>> getIssuerTicker(string queryType, IssuerTickerViewModel dataModel);

        MSGReturnModel saveIssuerTicker(string actionType, IssuerTickerViewModel dataModel);

        MSGReturnModel deleteIssuerTicker(string issuer);

        Tuple<bool, List<IssuerRatingViewModel>> getIssuerRating(string queryType, IssuerRatingViewModel dataModel);

        MSGReturnModel saveIssuerRating(string actionType, IssuerRatingViewModel dataModel);

        MSGReturnModel deleteIssuerRating(string issuer);

        Tuple<bool, List<GuarantorRatingViewModel>> getGuarantorRating(string queryType, GuarantorRatingViewModel dataModel);

        MSGReturnModel saveGuarantorRating(string actionType, GuarantorRatingViewModel dataModel);

        MSGReturnModel deleteGuarantorRating(string issuer);

        Tuple<bool, List<BondRatingViewModel>> getBondRating(string queryType, BondRatingViewModel dataModel);

        MSGReturnModel saveBondRating(string actionType, BondRatingViewModel dataModel);

        MSGReturnModel deleteBondRating(string bondNumber);
    }
}