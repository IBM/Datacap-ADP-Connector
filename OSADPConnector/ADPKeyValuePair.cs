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
    class ADPKeyValuePair : IComparable
    {
        public const String ADPKeyClassConfidence = "ADPKeyClassConfidence";
        public const String ADPKeyClass = "ADPKeyClassName";
        public const String ADPLineItemID = "ADPLineItemID";
        public const String ADPSeqLineItemID = "ADPSeqLineItemID";
        public const String ADPSensitivity = "ADPSensitivity";
        public ADPKeyValuePair()
        {
            this.nested = new List<ADPKeyValuePair>();
        }
        public List<ADPKeyValuePair> nested { get; set; }
        public String key { get; set; }
        public String value { get; set; }
        public int keyX { get; set; }
        public int keyY { get; set; }
        public int keyWidth { get; set; }
        public int keyHeight { get; set; }
        public int valueX { get; set; }
        public int valueY { get; set; }
        public int valueWidth { get; set; }
        public int valueHeight { get; set; }
        public String keyClassID { get; set; }
        public String keyClass { get; set; }
        public String keyClassConfidence { get; set; }
        public String kvpID { get; set; }
        public String originalKey { get; set; }
        public String originalValue { get; set; }
        public int confidence { get; set; }
        public Boolean haveLineItem { get; set; }
        public int LineItemID { get; set; }
        public int SeqLineItemID { get; set; }
        public Boolean sensitivity { get; set; }
        public Boolean added { get; set; }
        public int CompareTo(object obj)
        {
            ADPKeyValuePair other = (ADPKeyValuePair)obj;
            if ((this.keyClassConfidence != null) && (other.keyClassConfidence != null))
            {
                if (this.keyClassConfidence.Trim().ToLower().Equals("high") && !other.keyClassConfidence.Trim().ToLower().Equals("high"))
                {
                    return -1;
                }
                else if (other.keyClassConfidence.Trim().ToLower().Equals("high") && !this.keyClassConfidence.Trim().ToLower().Equals("high"))
                {
                    return 1;
                }
                else if (this.keyClassConfidence.Trim().ToLower().Equals("medium") && !other.keyClassConfidence.Trim().ToLower().Equals("medium"))
                {
                    return -1;
                }
                else if (other.keyClassConfidence.Trim().ToLower().Equals("medium") && !this.keyClassConfidence.Trim().ToLower().Equals("medium"))
                {
                    return 1;
                }
            }
            if (this.valueY < other.valueY)
            {
                return -1;
            }
            else if (this.valueY > other.valueY)
            {
                return 1;
            }
            else if (this.valueX < other.valueX)
            {
                return -1;
            }
            else if (this.valueX > other.valueX)
            {
                return 1;
            }
            else if (this.valueHeight < other.valueHeight)
            {
                return -1;
            }
            else if (this.valueHeight > other.valueHeight)
            {
                return 1;
            }
            else if (this.valueWidth < other.valueWidth)
            {
                return -1;
            }
            else if (this.valueWidth > other.valueWidth)
            {
                return 1;
            }
            else if (this.keyY < other.keyY)
            {
                return -1;
            }
            else if (this.keyY > other.keyY)
            {
                return 1;
            }
            else if (this.keyX < other.keyX)
            {
                return -1;
            }
            else if (this.keyX > other.keyX)
            {
                return 1;
            }
            else if (this.keyHeight < other.keyHeight)
            {
                return -1;
            }
            else if (this.keyHeight > other.keyHeight)
            {
                return 1;
            }
            else if (this.keyWidth < other.keyWidth)
            {
                return -1;
            }
            else if (this.keyWidth > other.keyWidth)
            {
                return 1;
            }
            return 0;
        }
    }
}
