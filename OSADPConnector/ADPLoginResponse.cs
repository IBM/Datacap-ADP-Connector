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
    /*
    Example response from ADP login with CP4BA 21.0.3
            {"username":"...userid...",
            "role":"Admin",
            "permissions":["can_work_with_ba_automations","can_administrate_business_teams","administrator","can_provision"],
            "groups":[10000],
            "sub":"...sub...",
            "iss":"KNOXSSO",
            "aud":"DSX",
            "uid":"1000331002",
            "authenticator":"external",
            "iam":{"accessToken":"37f56...596e4d"},
            "display_name":"cp4badmin",
            "accessToken":"eyJhb...mSiUNw",
            "_messageCode_":"success",
            "message":"success"}

    */
    public class ADPLoginResponse
    {
        /// <summary>
        /// This is the accessToken in the response
        /// </summary>
        public String accessToken { get; set; }

        /// <summary>
        /// This is the access_token in the response
        /// </summary>
        public String access_token { get; set; }

        public String getAccessToken()
        {
            if ((accessToken != null) && (accessToken.Trim().Length > 0))
            {
                return accessToken;
            }
            if ((access_token != null) && (access_token.Trim().Length > 0))
            {
                return access_token;
            }
            return null;
        }

        public override string ToString()
        {
            StringBuilder response = new StringBuilder();
            response.Append("{").Append(Environment.NewLine);
            response.Append("\"access token\": ").Append(getAccessToken()).Append(",").Append(Environment.NewLine);
            response.Append("}");
            return response.ToString();
        }

    }
}
