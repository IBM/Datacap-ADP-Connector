//
// © Copyright IBM Corp. 1994, 2023 All Rights Reserved
//
// Created by Scott Sumner-Moore, 2023
//

using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OSADPConnector
{
    public class ADPConnectInfo
    {
        /// <summary>
        /// This is the HttpClient that was connected
        /// </summary>
        public HttpClient client { get; set; }

        /// <summary>
        /// This is the Login Response
        /// </summary>
        public ADPLoginResponse adpLoginResponse { get; set; }

    }
}
