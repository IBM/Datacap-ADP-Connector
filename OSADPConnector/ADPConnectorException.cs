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
    public class ADPConnectorException : Exception
    {

        public ADPConnectorException()
        {

        }

        public ADPConnectorException(string message)
            : base(message)
        {
        }

        public ADPConnectorException(string message, Exception inner)
            : base(message, inner)
        {
        }



    }
}
