﻿using System;
using System.Collections.Generic;
using System.Text;

namespace WebScrapping
{
    public class ObjectModel
    {
        public string AuctionDate { get; set; }
        public string SecurityType { get; set; }
        public string Tenor { get; set; }
        public string AuctionNo { get; set; }
        public string Auction { get; set; }
        public string MaturityDate { get; set; }
        public string TotalSubscription { get; set; }
        public string TotalSuccessful { get; set; }
        public string RangeBid { get; set; }
        public string SuccessfulBidRates { get; set; }
        public string Description { get; set; }
        public string Rate { get; set; }
        public string TrueYield { get; set; }
        public string AmountOffered { get; set; }
    }
}
