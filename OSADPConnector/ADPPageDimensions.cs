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
    class ADPPageDimensions
    {
        public int PageHeight { get; set; }
        public int PageWidth { get; set; }
        public int dpiX { get; set; }
        public int dpiY { get; set; }
        public float PageOCRCOnfidence { get; set; }
    }
}
