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
    class ADPProcessingStatus
    {
        /*
        Sample data:
        {
            "status": {
                "code": 200,
                "messageId": "CIWCA50000",
                "message": "Success"
            },
            "result": [
                {
                    "status": {
                        "code": 200,
                        "messageId": "CIWCA11107",
                        "message": "Successfully retrieved the content analyzer details"
                    },
                    "data": {
                        "analyzerId": "195...479",
                        "uniqueId": "",
                        "creationDate": "2021-08-31T22:59:30.262Z",
                        "fileName": "TM000001.tif",
                        "numPages": 1,
                        "statusDetails": [
                            {
                                "type": "JSON",
                                "status": "InProgress",
                                "startTime": "2021-08-31T22:59:30.262Z",
                                "completedPages": 0,
                                "progress": 8
                            }
                        ]
                    }
                }
            ]
        }

        */
        public class Status
        {
            public int code { get; set; }
            public string messageId { get; set; }
            public string message { get; set; }
        }

        public class StatusDetail
        {
            public string type { get; set; }
            public string status { get; set; }
            public DateTime startTime { get; set; }
            public int completedPages { get; set; }
            public int progress { get; set; }
        }

        public class Data
        {
            public string analyzerId { get; set; }
            public string uniqueId { get; set; }
            public DateTime creationDate { get; set; }
            public string fileName { get; set; }
            public int numPages { get; set; }
            public List<StatusDetail> statusDetails { get; set; }
        }

        public class Result
        {
            public Status status { get; set; }
            public Data data { get; set; }
        }

        public Status status { get; set; }
        public List<Result> result { get; set; }

    }
}
