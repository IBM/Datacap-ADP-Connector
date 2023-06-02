//
// © Copyright IBM Corp. 1994, 2023 All Rights Reserved
//
// Created by Scott Sumner-Moore, 2023
//


using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OSADPConnector
{
    class ADPRanking
    {
        public String keyClassID { get; set; }
        public String keyClassName { get; set; }
        public String keyClassType { get; set; }
        public List<ADPRank> kvpRankedList { get; set; }
    }
}
