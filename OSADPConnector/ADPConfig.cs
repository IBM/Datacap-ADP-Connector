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
    Example adp.json configuration for CP4BA 21.0.3
    {
    	"zen_base_url": "https:\/\/cpd-adp....containers.appdomain.cloud",
    	"login_target": "\/v1\/preauth\/validateAuth",
    	"verify_token_target": "\/usermgmt\/v1\/user\/currentUserInfo",
    	"adp_project_id": "aed...250",
    	"analyze_target": "\/adp\/aca\/v1\/projects\/[[adp_project_id]]\/analyzers",
    	"json_options": "ocr,dc,kvp,sn,hr,th,mt,ai,ds",
    	"multiThreading": "false",
    	"output_directory_path": "C:\\Temp\\StandaloneADP",
    	"output_options": "json",
    	"ssl_verification": "false",
    	"ums_password": "...password...",
    	"ums_username": "...userid...",
        "timeout_in_minutes": "5",
        "document_class": "Bill of Lading"
    }
    */
    public class ADPConfig
    {
        public const String ADPDocType = "ADPDocType";

        /// <summary>
        /// This is the zen_base_url for the service
        /// </summary>
        public String zen_base_url { get; set; }

        /// <summary>
        /// This is the ums_base_url for the service
        /// </summary>
        public String ums_base_url { get; set; }

        /// <summary>
        /// This is the aca_base_url for the service
        /// </summary>
        public String aca_base_url { get; set; }

        /// <summary>
        /// This is the login target for the service (the part after the ums_base_url or zen_base_url)
        /// </summary>
        public String login_target { get; set; }

        /// <summary>
        /// This is the verify token target for the service (the part after zen_base_url)
        /// </summary>
        public String verify_token_target { get; set; }

        /// <summary>
        /// This is the target to analyze documents (the part after the aca_base_url or zen_base_url)
        /// The adp_project_id is embedded in this string
        /// </summary>
        public String analyze_target { get; set; }



        /// <summary>
        /// This is the ums_username that should be used to connect to the ADP Server
        /// </summary>
        public String ums_username { get; set; }

        /// <summary>
        /// This is the ums_password for the user to connect to the ADP Server
        /// </summary>
        public String ums_password { get; set; }

        /// <summary>
        /// This is the zen_username that should be used to connect to the ADP Server
        /// </summary>
        public String zen_username { get; set; }

        /// <summary>
        /// This is the zen_password for the user to connect to the ADP Server
        /// </summary>
        public String zen_password { get; set; }

        /// <summary>
        /// This is the ADP project id 
        /// </summary>
        public String adp_project_id { get; set; }

        /// <summary>
        /// This is the client id for connecting to the ADP server
        /// </summary>
        public String client_id { get; set; }

        /// <summary>
        /// This is the client secret for connecting to the ADP Server
        /// </summary>
        public String client_secret { get; set; }

        /// <summary>
        /// This is the output directory path for writing the json returned from the ADP Server
        /// </summary>
        //public String output_directory_path { get; set; }

        /// <summary>
        /// This is the output options for the call to the ADP Server
        /// </summary>
        public String output_options { get; set; }

        /// <summary>
        /// This is the json options for the processing on the ADP Server
        /// </summary>
        public String json_options { get; set; }

        /// <summary>
        /// If this is set to true, then each page will be sent to the ADP Service with a different thread
        /// </summary>
        public String multiThreading { get; set; }

        /// <summary>
        /// If this is set, it will be passed to ADP as the assigned document class to use for field extraction
        /// </summary>
        public String docClass { get; set; }

        /// <summary>
        /// The suffix to use on all ADP fields when they are put onto the Datacap page. If missing, defaults to "_ADP".
        /// Can be set to empty string for no suffix. Note that if there are multiple fields with the same ADP key class,
        /// the field on the Datacap page will be key-class-name followed by the field_suffix, followed by _nn.
        /// For example, "Bill date_ADP_0", "Bill date_ADP_1", etc.
        /// </summary>
        public String field_suffix { get; set; }

        /// <summary>
        /// How long to wait, in minutes, for ADP to complete processing a page
        /// </summary>
        public String timeout_in_minutes { get; set; }

        /// <summary>
        /// How long to wait, in minutes, for ADP to complete processing a page
        /// </summary>
        public String use_all_pages { get; set; }

        public int getTimeoutInMinutes()
        {
            int timeout = 5;
            if ((this.timeout_in_minutes == null) || (this.timeout_in_minutes.Trim().Length == 0) || !Int32.TryParse(this.timeout_in_minutes, out timeout))
            {
                return timeout; // Default to 5 minute timeout
            }
            return timeout;
        }

        public String getAnalyzeURL()
        {
            if ((zen_base_url != null) && (zen_base_url.Trim().Length > 0))
            {
                return zen_base_url + analyze_target;
            }
            else if ((aca_base_url != null) && (aca_base_url.Trim().Length > 0))
            {
                return aca_base_url + analyze_target;
            }
            else
            {
                throw new ADPConnectorException("error in ADP configuration file -- need either zen_base_url or ums_base_url");
            }
        }

        public String getLoginURL()
        {
            if ((zen_base_url != null) && (zen_base_url.Trim().Length > 0))
            {
                return zen_base_url + login_target;
            }
            else if ((ums_base_url != null) && (ums_base_url.Trim().Length > 0))
            {
                return ums_base_url + login_target;
            }
            else
            {
                throw new ADPConnectorException("error in ADP configuration file -- need either zen_base_url or ums_base_url");
            }

        }

    }
}
