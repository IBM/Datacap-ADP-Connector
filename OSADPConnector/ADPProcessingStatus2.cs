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
    class ADPProcessingStatus2
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
            public float progress { get; set; }
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

        public ADPProcessingStatus cloneToADPProcessingStatus()
        {
            ADPProcessingStatus s = new ADPProcessingStatus();
            s.status = new ADPProcessingStatus.Status();
            s.status.code = this.status.code;
            s.status.message = this.status.message;
            s.status.messageId = this.status.messageId;
            s.result = new List<ADPProcessingStatus.Result>();
            foreach (Result thisResult in this.result)
            {
                ADPProcessingStatus.Result sResult = new ADPProcessingStatus.Result();
                sResult.status = new ADPProcessingStatus.Status();
                sResult.status.code = thisResult.status.code;
                sResult.status.message = thisResult.status.message;
                sResult.status.messageId = thisResult.status.messageId;
                sResult.data = new ADPProcessingStatus.Data();
                sResult.data.analyzerId = thisResult.data.analyzerId;
                sResult.data.creationDate = thisResult.data.creationDate;
                sResult.data.fileName = thisResult.data.fileName;
                sResult.data.numPages = thisResult.data.numPages;
                sResult.data.uniqueId = thisResult.data.uniqueId;
                sResult.data.statusDetails = new List<ADPProcessingStatus.StatusDetail>();
                foreach (StatusDetail thisStatusDetail in thisResult.data.statusDetails)
                {
                    ADPProcessingStatus.StatusDetail sStatusDetail = new ADPProcessingStatus.StatusDetail();
                    sStatusDetail.completedPages = thisStatusDetail.completedPages;
                    sStatusDetail.progress = (int)thisStatusDetail.progress;
                    sStatusDetail.startTime = thisStatusDetail.startTime;
                    sStatusDetail.status = thisStatusDetail.status;
                    sStatusDetail.type = thisStatusDetail.type;
                    sResult.data.statusDetails.Add(sStatusDetail);
                }
                s.result.Add(sResult);
            }
            return s;
        }

    }
}
