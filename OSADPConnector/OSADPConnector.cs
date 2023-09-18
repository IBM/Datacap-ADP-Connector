//
// © Copyright IBM Corp. 1994, 2023 All Rights Reserved
//
// Created by Scott Sumner-Moore, 2023
//
// This is a .NET action for IBM Datacap that provides a bidirectional integration with ADP.
// The compliled DLL needs to be placed into the RRS or C:\Datacap\[yourapp]\dco_[yourapp]\rules directory.
// The DLL does not need to be registered.  
// Datacap studio will find the RRX file that is embedded in the DLL, you do not need to place the RRX in the RRS directory.
// Newtonsoft.Json.dll is referenced by this code; that DLL should be added to C:\RRS or C:\Datacap\[yourapp]\dco_[yourapp]\rules so it can be found at runtime.
// If Datacap references are not found at compile time, add a reference path of C:\Datacap\DCShared\NET to the project to locate the DLLs while building.
// This code has been tested with IBM Datacap 9.1.8 and ADP 21.0.3, 22.0.1 and 22.0.2.

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.Serialization.Json;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Newtonsoft.Json;

namespace OSADPConnector
{
    public class OSADPConnector // This class must be a base class for .NET 4.0 Actions.
    {
        #region ExpectedByRRS
        /// <summary/>
        ~OSADPConnector()
        {
            DatacapRRCleanupTime = true;
        }

        /// <summary>
        /// Cleanup: This property is set right before the object is released by RRS
        /// The Dispose method is not called by RRS.
        /// </summary>
        public bool DatacapRRCleanupTime
        {
            set
            {
                if (value)
                {
                    CleanUp();
                    CurrentDCO = null;
                    DCO = null;
                    RRLog = null;
                    RRState = null;
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                }
            }
        }

        protected PILOTCTRLLib.IBPilotCtrl BatchPilot = null;
        public PILOTCTRLLib.IBPilotCtrl DatacapRRBatchPilot { set { this.BatchPilot = value; GC.Collect(); GC.WaitForPendingFinalizers(); } get { return this.BatchPilot; } }

        protected TDCOLib.IDCO DCO = null;
        /// <summary/>
        public TDCOLib.IDCO DatacapRRDCO
        {
            get { return this.DCO; }
            set
            {
                DCO = value;
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        protected dcrroLib.IRRState RRState = null;
        /// <summary/>
        public dcrroLib.IRRState DatacapRRState
        {
            get { return this.RRState; }
            set
            {
                RRState = value;
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        public TDCOLib.IDCO CurrentDCO = null;
        /// <summary/>
        public TDCOLib.IDCO DatacapRRCurrentDCO
        {
            get { return this.CurrentDCO; }
            set
            {
                CurrentDCO = value;
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        public dclogXLib.IDCLog RRLog = null;
        /// <summary/>
        public dclogXLib.IDCLog DatacapRRLog
        {
            get { return this.RRLog; }
            set
            {
                RRLog = value;
                LogAssemblyVersion();
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        #endregion

        #region CommonActions

        void OutputToLog(int nLevel, string strMessage)
        {
            if (null == RRLog)
                return;
            RRLog.WriteEx(nLevel, strMessage);
        }

        public void WriteLog(string sMessage)
        {
            OutputToLog(5, sMessage);
        }

        private bool versionWasLogged = false;

        // Log the version of the library that was running to help with diagnosis.
        // Hooked this method to be called after the log object is assigned.  Also put in
        // a check that this action runs only once, just in case it gets called multiple times.
        protected bool LogAssemblyVersion()
        {
            try
            {
                if (versionWasLogged == false)
                {
                    FileVersionInfo fv = System.Diagnostics.FileVersionInfo.GetVersionInfo(Assembly.GetExecutingAssembly().Location);
                    WriteLog(Assembly.GetExecutingAssembly().Location +
                             ". AssemblyVersion: " + Assembly.GetExecutingAssembly().GetName().Version.ToString() +
                             ". AssemblyFileVersion: " + fv.FileVersion.ToString() + ".");
                    versionWasLogged = true;
                }
            }
            catch (Exception ex)
            {
                WriteLog("Version logging exception: " + ex.Message);
            }

            // We can always return true.  If getting the version fails, we can try to continue anyway.
            return true;
        }

        #endregion


        // implementation of the Dispose method to release managed resources
        // There is no guarentee that dispose will be called.  Also note, class distructors are also not called.  CleanupTime is called by RRS.        
        public void Dispose()
        {
            CleanUp();
        }

        /// <summary>
        /// Everthing that should be cleaned up on exit
        /// It is recommended to avoid logging during cleanup.
        /// </summary>
        protected void CleanUp()
        {
            try
            {
                // Cleanup and release any allocated objects here. This will be called before the DLL is released.
            }
            catch { } // Ignore any errors.
        }

        struct Level
        {
            internal const int Batch = 0;
            internal const int Document = 1;
            internal const int Page = 2;
            internal const int Field = 3;
            internal const int Character = 4;
        }

        struct Status
        {
            internal const int Hidden = -1;
            internal const int OK = 0;
            internal const int Fail = 1;
            internal const int Over = 3;
            internal const int RescanPage = 70;
            internal const int VerificationFailed = 71;
            internal const int PageOnHold = 72;
            internal const int PageOverridden = 73;
            internal const int NoData = 74;
            internal const int DeletedPage = 75;
            internal const int ExportComplete = 76;
            internal const int DeleteApproved = 77;
            internal const int ReviewPage = 79;
            internal const int DeletedDoc = 128;
        }

        public static int LOGGING_INFO = 3;
        public static int LOGGING_DEBUG = 1;
        public static int LOGGING_ERROR = 5;

        private TDCOLib.IDCO oBatch;
        public TDCOLib.IDCO oPage; // TODO
        public dcSmart.SmartNav smartNav = null;
        private Utils utils;
        private String keyFile = null;
        private Boolean standalone = false;
        private int adpFieldsAction;
        private int pollingIntervalInSeconds = 10;
        private int maxThreads = 3;

        // What to do with ADP fields
        public const int ADPFIELDS_KEEPALLWITHKEYCLASS = 1;
        public const int ADPFIELDS_KEEPALL = 2;
        public const int ADPFIELDS_KEEPSINGLEBEST = 3;
        public const int ADPFIELDS_KEEPSINGLEBESTWITHKEYCLASS = 4;

        // What to do with ADP entities after resolving ADP vs Accelerator
        public const int ADP_KEEPALL = 1;
        public const int ADP_KEEPHIGH = 2;
        public const int ADP_KEEPMEDIUM = 3;
        public const int ADP_DELETEALL = 4;

        private int convertConfidenceString(String confidenceAction)
        {
            int adpConfidenceAction = ADP_KEEPALL;
            if ((confidenceAction != null) && (confidenceAction.Trim().Length > 0))
            {
                if (confidenceAction.Trim().ToLower().Equals("keepall"))
                {
                    adpConfidenceAction = ADP_KEEPALL;
                }
                else if (confidenceAction.Trim().ToLower().Equals("keephigh"))
                {
                    adpConfidenceAction = ADP_KEEPHIGH;
                }
                else if (confidenceAction.Trim().ToLower().Equals("keepmedium"))
                {
                    adpConfidenceAction = ADP_KEEPMEDIUM;
                }
                else if (confidenceAction.Trim().ToLower().Equals("deleteall"))
                {
                    adpConfidenceAction = ADP_DELETEALL;
                }
                else
                {
                    RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "Parameter for confidenceAction is not valid: " + confidenceAction);
                }
            }
            return adpConfidenceAction;
        }

        private int convertFieldsString(String extraFieldsAction)
        {
            int adpFieldsAction = ADPFIELDS_KEEPALL;
            if ((extraFieldsAction != null) && (extraFieldsAction.Trim().Length > 0))
            {
                if (extraFieldsAction.Trim().ToLower().Equals("keepallwithkeyclass"))
                {
                    adpFieldsAction = ADPFIELDS_KEEPALLWITHKEYCLASS;
                }
                else if (extraFieldsAction.Trim().ToLower().Equals("keepsinglebest"))
                {
                    adpFieldsAction = ADPFIELDS_KEEPSINGLEBEST;
                }
                else if (extraFieldsAction.Trim().ToLower().Equals("keepsinglebestwithkeyclass"))
                {
                    adpFieldsAction = ADPFIELDS_KEEPSINGLEBESTWITHKEYCLASS;
                }
                else if (extraFieldsAction.Trim().ToLower().Equals("keepall"))
                {
                    adpFieldsAction = ADPFIELDS_KEEPALL;
                }
                else
                {
                    RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "Parameter for extraFieldsAction is not valid: " + extraFieldsAction);
                }
            }
            return adpFieldsAction;
        }

        private Int64 processADP(String pageID, TDCOLib.IDCO oPage, ADPConfig adpConfig)
        {
            Int64 startTime = DateTimeOffset.Now.ToUnixTimeMilliseconds();
            try
            {
                String jsonResults = null;

                if (true)
                {

                    RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "RestServicesAction: processADP: oPage: " + oPage);
                    RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "RestServicesAction: processADP: BatchPilot: " + BatchPilot);

                    String tifFile = oPage.ImageName;
                    String ocrJsonFile = BatchPilot.BatchDir + "\\" + pageID + "_adp.json";

                    RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "RestServicesAction: processADP: Image file is: " + tifFile);
                    RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "RestServicesAction: processADP: ADP JSON Output file will be : " + ocrJsonFile);

                    int xDpi = 200;
                    int yDpi = 200;
                    int width = 0;
                    int height = 0;
                    if (!tifFile.EndsWith(".pdf", true, System.Globalization.CultureInfo.InvariantCulture))
                    {
                        using (FileStream stream = new FileStream(tifFile, FileMode.Open, FileAccess.Read))
                        {
                            using (Image tif = Image.FromStream(stream, false, false))
                            {
                                xDpi = Convert.ToInt32(tif.HorizontalResolution);
                                yDpi = Convert.ToInt32(tif.VerticalResolution);
                                width = Convert.ToInt32(tif.PhysicalDimension.Width);
                                height = Convert.ToInt32(tif.PhysicalDimension.Height);
                            }
                            stream.Close();
                        }
                    }

                    Boolean done = false;
                    Boolean alreadyRetried = false;
                    while (!done)
                    {
                        // Connect to Server
                        ADPConnectInfo adpCI = connectToADP(adpConfig, RRLog);
                        HttpClient client = adpCI.client;
                        ADPLoginResponse adpLR = adpCI.adpLoginResponse;

                        // Send Page to Service
                        String url = adpConfig.getAnalyzeURL(); // .aca_base_url + adpConfig.analyze_target;
                        RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "processADP: analyze url: " + url);
                        String analyzerID = uploadToADP(client, url, adpConfig, adpLR, xDpi, File.ReadAllBytes(tifFile), tifFile, RRLog);
                        RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "processADP: analyzerID: " + analyzerID);

                        // Wait for result from ADP]
                        Boolean retry = false;
                        Boolean success = false;
                        waitUntilDoneInADP(client, url + "/" + analyzerID, adpLR, RRLog, adpConfig.getTimeoutInMinutes(), out retry, out success);
                        //Can this be done in a way that yields time back to RuleRunner? Would probably have to wait and loop in Datacap actions
                        if (!retry) // if we haven't timed out the connection, keep going: get the results, etc.
                        {
                            done = true;
                            if (success)
                            {
                                // Get the ADP JSON data
                                jsonResults = getResultFromADP(client, url + "/" + analyzerID + "/json", adpLR, RRLog);

                                // Write data to json file
                                File.WriteAllText(ocrJsonFile, jsonResults);

                                // Delete the document from ADP
                                deleteFromADP(client, url + "/" + analyzerID, adpLR, RRLog);

                                // Dispose of the session
                                client.Dispose();
                            }
                            else
                            {
                                RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "processADP: processing timed out -- failing this document");

                                // Dispose of the session
                                client.Dispose();

                                throw new ADPConnectorException("ADP processing timed out");
                            }
                        }
                        else
                        {
                            if (alreadyRetried)
                            {
                                done = true;
                                RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "processADP: connection timed out twice -- failing this document");
                                throw new ADPConnectorException("Timed out ADP connection twice -- no more!");
                            }
                            else
                            {
                                alreadyRetried = true;
                                // If we timed out the connection, loop back around and try again
                                RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "processADP: connection timed out -- trying again");
                            }
                        }
                    }

                    String smartNavDocClass = null;
                    if ((adpConfig != null) && (adpConfig.docClass != null))
                    {
                        smartNavDocClass = smartNav.MetaWord(adpConfig.docClass);
                    }
                    if ((smartNavDocClass != null) && (smartNavDocClass.Trim().Length > 0))
                    {
                        oPage.Type = smartNavDocClass;
                    }
                    else
                    {
                        int docClassVarNum = oPage.FindVariable("ADPDocType");
                        if (docClassVarNum >= 0)
                        {
                            String docClassFromADP = oPage.GetVariableValue(docClassVarNum);
                            oPage.Type = docClassFromADP;
                        }
                    }

                }
                else
                {
                    // For debugging/testing without actually calling ADP -- reads ADP response from directory
                    RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "before reading adp json: " + BatchPilot.BatchDir + "\\" + pageID + "_adp.json");
                    jsonResults = File.ReadAllText(BatchPilot.BatchDir + "\\" + pageID + "_adp.json");
                    RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "after reading adp json");
                }

                processADPJson(jsonResults, adpConfig.field_suffix, oPage, pageID, 0); // Single page for this call

            }
            catch (Exception e)
            {
                RRLog.WriteEx(OSADPConnector.LOGGING_ERROR, "error : " + e.Message);
                RRLog.WriteEx(OSADPConnector.LOGGING_ERROR, e.ToString());
                String className = this.GetType().Name;
                String methodName = System.Reflection.MethodInfo.GetCurrentMethod().Name;
                RRLog.WriteEx(OSADPConnector.LOGGING_ERROR, "Class " + className + ", method " + methodName);
                RRLog.WriteEx(OSADPConnector.LOGGING_ERROR, "Stack Trace: " + e.StackTrace);
            }
            // Return the time in ms to complete the task
            return DateTimeOffset.Now.ToUnixTimeMilliseconds() - startTime;
        }

        private ADPConnectInfo connectToADP(ADPConfig config, dclogXLib.IDCLog RRLog)
        {
            ADPConnectInfo adpCI = new ADPConnectInfo();
            HttpClient httpClient = null;
            ADPLoginResponse adpLR = new ADPLoginResponse();

            ServicePointManager.ServerCertificateValidationCallback +=
                (sender, cert, chain, sslPolicyErrors) => { return true; };

            HttpClientHandler httpClientHandler = new HttpClientHandler()
            {
                Credentials = new NetworkCredential(config.zen_username.Trim(), config.zen_password.Trim()),
            };

            httpClient = new HttpClient(httpClientHandler);
            httpClient.Timeout = TimeSpan.FromMinutes(10); // Some really huge documents were taking longer than the default of 100 seconds

            RRLog.WriteEx(OSADPConnector.LOGGING_ERROR, "RestServices: connectToADP2103: after default request headers");

            HttpRequestMessage httpRequestMessage = new HttpRequestMessage();
            httpRequestMessage.Method = HttpMethod.Get;
            RRLog.WriteEx(OSADPConnector.LOGGING_ERROR, "RestServices: connectToADP2103: zen_base_url: " + config.zen_base_url);
            RRLog.WriteEx(OSADPConnector.LOGGING_ERROR, "RestServices: connectToADP2103: login_target: " + config.login_target);
            httpRequestMessage.RequestUri = new Uri(config.getLoginURL());
            String encoding = config.zen_username + ":" + config.zen_password;
            byte[] b64 = System.Text.Encoding.UTF8.GetBytes(encoding);
            encoding = System.Convert.ToBase64String(b64);
            httpRequestMessage.Headers.Add("authorization", "Basic " + encoding);
            httpRequestMessage.Headers.Add("'Content-Type", "application/x-www-form-urlencoded");

            RRLog.WriteEx(OSADPConnector.LOGGING_ERROR, "RestServices: connectToADP2103: httpRequestMessage.Headers: " + httpRequestMessage.Headers);
            RRLog.WriteEx(OSADPConnector.LOGGING_ERROR, "RestServices: connectToADP2103: httpRequestMessage.RequestUri: " + httpRequestMessage.RequestUri);

            HttpResponseMessage response = httpClient.SendAsync(httpRequestMessage).Result;

            /*
            {"username":"cp4badmin",
            "role":"Admin",
            "permissions":["can_work_with_ba_automations","can_administrate_business_teams","administrator","can_provision"],
            "groups":[10000],
            "sub":"cp4badmin",
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

            byte[] bytes = null;
            if (response.IsSuccessStatusCode)
            {
                bytes = response.Content.ReadAsByteArrayAsync().Result;
                RRLog.WriteEx(OSADPConnector.LOGGING_ERROR, "RestServices: connectToADP2103: response length: " + bytes.Length);
                String respAsString = response.Content.ReadAsStringAsync().Result;
                RRLog.WriteEx(OSADPConnector.LOGGING_ERROR, "RestServices: connectToADP2103: ADP Login Response: " + respAsString);
                if (bytes != null)
                {
                    DataContractJsonSerializer ser = new DataContractJsonSerializer(typeof(ADPLoginResponse));
                    StreamReader sr = new StreamReader(new MemoryStream(bytes));
                    adpLR = (ADPLoginResponse)ser.ReadObject(sr.BaseStream);
                    sr.Close();
                }

                //RRLog.WriteEx(OSADPConnector.LOGGING_ERROR, "RestServices: connectToADP2103: bytes length: " + bytes.Length);
                //String respAsString = response.Content.ReadAsStringAsync().Result;
                //RRLog.WriteEx(OSADPConnector.LOGGING_ERROR, "RestServices: connectToADP2103: ADP Login Response: " + respAsString);
                if (RRLog != null)
                {
                    RRLog.WriteEx(OSADPConnector.LOGGING_ERROR, "RestServices: connectToADP21031: HTTP Code: " + response.StatusCode);
                    RRLog.WriteEx(OSADPConnector.LOGGING_ERROR, "RestServices: connectToADP21031: Reason: " + response.ReasonPhrase);
                    if (response.RequestMessage != null)
                    {
                        RRLog.WriteEx(OSADPConnector.LOGGING_ERROR, "RestServices: connectToADP21031: Request Message: " + response.RequestMessage.ToString());
                    }
                    RRLog.WriteEx(OSADPConnector.LOGGING_ERROR, "RestServices: connectToADP21031: Response Message: " + response.ToString());
                }
                else
                {
                    Console.WriteLine("RestServices: connectToADP21032: HTTP Code: " + response.StatusCode);
                    Console.WriteLine("RestServices: connectToADP21032: Reason: " + response.ReasonPhrase);
                    if (response.RequestMessage != null)
                    {
                        Console.WriteLine("RestServices: connectToADP21032: Request Message: " + response.RequestMessage.ToString());
                    }
                    Console.WriteLine("RestServices: connectToADP21032: Response Message: " + response.ToString());
                }
            }
            else
            {
                // Some sort of error with first send, report the message in the log
                if (RRLog != null)
                {
                    RRLog.WriteEx(OSADPConnector.LOGGING_ERROR, "RestServices: connectToADP2103: HTTP Code: " + response.StatusCode);
                    RRLog.WriteEx(OSADPConnector.LOGGING_ERROR, "RestServices: connectToADP2103: Reason: " + response.ReasonPhrase);
                    RRLog.WriteEx(OSADPConnector.LOGGING_ERROR, "RestServices: connectToADP2103: Request Message: " + response.RequestMessage.ToString());
                    RRLog.WriteEx(OSADPConnector.LOGGING_ERROR, "RestServices: connectToADP2103: Response Message: " + response.ToString());

                }
                // Set up for simple authorization in all calls
                httpClient.DefaultRequestHeaders.Add("authorization", "Basic " + encoding);
                /*
                                // Try again
                                response = httpClient.SendAsync(httpRequestMessage).Result;

                                RRLog.WriteEx(OSADPConnector.LOGGING_ERROR, "RestServices: connectToADP21031: after httpClient.SendAsync");
                                if (response.IsSuccessStatusCode)
                                {
                                    bytes = response.Content.ReadAsByteArrayAsync().Result;
                                    if (bytes != null)
                                    {
                                        DataContractJsonSerializer ser = new DataContractJsonSerializer(typeof(ADPLoginResponse));
                                        StreamReader sr = new StreamReader(new MemoryStream(bytes));
                                        adpLR = (ADPLoginResponse)ser.ReadObject(sr.BaseStream);
                                        sr.Close();
                                    }
                                }
                                else
                                {
                                    if (RRLog != null)
                                    {
                                        RRLog.WriteEx(OSADPConnector.LOGGING_ERROR, "RestServices: connectToADP2103: HTTP Code: " + response.StatusCode);
                                        RRLog.WriteEx(OSADPConnector.LOGGING_ERROR, "RestServices: connectToADP2103: Reason: " + response.ReasonPhrase);
                                        if (response.RequestMessage != null)
                                        {
                                            RRLog.WriteEx(OSADPConnector.LOGGING_ERROR, "RestServices: connectToADP21031: Request Message: " + response.RequestMessage.ToString());
                                        }
                                        RRLog.WriteEx(OSADPConnector.LOGGING_ERROR, "RestServices: connectToADP2103: Response Message: " + response.ToString());

                                    }
                                    //throw new ADPConnectorException(response.ReasonPhrase);
                                    adpLR = null;
                                }
                */
            }

            //String data = Encoding.UTF8.GetString(bytes);

            //// Write data to json file
            //File.WriteAllText(config.output_directory_path + "\\adpConnect.json", data);
            //if (adpLR != null)
            //{
            //    RRLog.WriteEx(OSADPConnector.LOGGING_ERROR, "RestServices: connectToADP2103: ADP Login Response: " + adpLR.ToString());
            //}

            adpCI.client = httpClient;
            adpCI.adpLoginResponse = adpLR;
            return adpCI;


        }


        public void sendToADP(String config)
        {
            RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "sendToADP: begin");
            ADPConfig adpConfig = null;
            //RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "sendToADP: About to load the ADP config string: " + config);
            adpConfig = Utils.loadADPConfig(config, RRLog);
            RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "sendToADP: adpConfig analyze_target: " + adpConfig.analyze_target);
            Boolean multiThread = false;
            if (adpConfig.multiThreading != null)
            {
                if (adpConfig.multiThreading.Trim().ToLower().Equals("true"))
                {
                    multiThread = true;
                }
            }


            // Loop through the pages and create a task for each page
            List<Task<Int64>> taskList = new List<Task<Int64>>();
            long singleThread = 0;
            if (this.CurrentDCO == null)
            {
                RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "sendToADP: How can this.CurrentDCO be null???");
            }
            if (this.CurrentDCO.ObjectType() == Level.Batch || this.CurrentDCO.ObjectType() == Level.Document)
            {
                if (this.CurrentDCO.ObjectType() == Level.Batch)
                {
                    oBatch = this.CurrentDCO;
                }
                if (oBatch == null)
                {
                    TDCOLib.IDCO parent = CurrentDCO.Parent();
                    while (parent.ObjectType() != Level.Batch)
                    {
                        parent = parent.Parent();
                    }
                    oBatch = parent;
                }
                /*
                                for (int i = 0; i < oBatch.NumOfChildren(); i++)
                                {
                                    TDCOLib.IDCO oPage = oBatch.GetChild(i);
                                    if (oPage.ObjectType() == Level.Page)
                                    {
                                        if (multiThread)
                                        {
                                            //Func<long> p = async () =>
                                            //{
                                            //    long multiThreadValue = await processADPAsync(oPage.ID, oPage, config);
                                            //    return multiThreadValue;
                                            //};
                                            taskList.Add(Task<long>.Factory.StartNew(() => processADP(oPage.ID, oPage, adpConfig))); // This will allow threading to occur
                                        }
                                        else
                                        {
                                            singleThread = processADP(oPage.ID, oPage, adpConfig);
                                        }
                                    }
                                }
                */
                if (multiThread)
                {
                    int currPage = 0;
                    int currNumThreads = 0;
                    int currWaitThread = 0;
                    Boolean done = false;
                    while (!done)
                    {
                        if (currPage < oBatch.NumOfChildren())
                        {
                            if (currNumThreads < maxThreads)
                            {
                                TDCOLib.IDCO oPage = oBatch.GetChild(currPage);
                                currPage++;
                                if (oPage.ObjectType() == Level.Page)
                                {
                                    taskList.Add(Task<long>.Factory.StartNew(() => processADP(oPage.ID, oPage, adpConfig))); // This will allow threading to occur
                                    currNumThreads++;
                                }
                            }
                            else
                            {
                                Double result = taskList[currWaitThread].Result;
                                RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "sendToADP: ADP: Took " + result + " to process page");
                                currWaitThread++;
                                currNumThreads--;
                            }
                        }
                        else
                        {
                            Double result = taskList[currWaitThread].Result;
                            RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "sendToADP: ADP: Took " + result + " to process page");
                            currWaitThread++;
                            currNumThreads--;
                        }
                        if ((currPage >= oBatch.NumOfChildren()) && (currNumThreads == 0))
                        {
                            done = true;
                        }
                    }
                }
                else
                {
                    for (int i = 0; i < oBatch.NumOfChildren(); i++)
                    {
                        TDCOLib.IDCO oPage = oBatch.GetChild(i);
                        if (oPage.ObjectType() == Level.Page)
                        {
                            singleThread = processADP(oPage.ID, oPage, adpConfig);
                        }
                    }
                }
            }
            else if (this.CurrentDCO.ObjectType() == Level.Page)
            {
                multiThread = false; // Not possible to multithread a single page
                RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "sendToADP: About to call processADP");
                singleThread = processADP(this.CurrentDCO.ID, this.CurrentDCO, adpConfig);
            }



            if (multiThread)
            {
                /*
                // Loop through the results, the output will be the time
                Double[] results = new Double[taskList.Count];
                for (int i = 0; i < taskList.Count; i++)
                {
                    results[i] = taskList[i].Result;
                    RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "sendToADP: ADP: Took " + results[i] + " to process page");
                }
                */
            }
            else
            {
                RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "sendToADP: ADP: Took " + singleThread + " to process all pages");
            }

            RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "sendToADP: end");
        }


        public async Task<bool> SendPageToADP(String zenBaseURL, String loginTarget, String zenUserName, String zenPassword, String analyzeTarget, String verifyTokenTarget, String adpProjectID,
                                      String fieldSuffix, String timeoutInMinutes, String jsonOptions, /*String outputDirectoryPath,*/ String outputOptions,
                                      String extraFieldsAction, String confidenceAction, String docClass, String stringPollingIntervalInSeconds, String stringMaxThreads)
        {
            try
            {
                DateTime compileTime = new DateTime(Builtin.CompileTime, DateTimeKind.Utc); // Ignore the supposed error -- this is resolved at compile time
                RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "***Start SendPageToADP - compile time " + compileTime);
                this.oPage = CurrentDCO;
                this.smartNav = new dcSmart.SmartNav(this);
                utils = new Utils(RRLog);
                RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "extraFieldsAction: " + extraFieldsAction + ", smartNav.MetaWord(extraFieldsAction): " + smartNav.MetaWord(extraFieldsAction));
                adpFieldsAction = convertFieldsString(smartNav.MetaWord(extraFieldsAction));
                int adpConfidenceAction = convertConfidenceString(smartNav.MetaWord(confidenceAction));
                string UseAllPages = smartNav.MetaWord("@P.UseAllPages");
                String sPollingIntervalInSeconds = smartNav.MetaWord(stringPollingIntervalInSeconds);
                if (!Int32.TryParse(sPollingIntervalInSeconds, out pollingIntervalInSeconds))
                {
                    pollingIntervalInSeconds = 10;
                }
                String sMaxThreads = smartNav.MetaWord(stringMaxThreads);
                if (!Int32.TryParse(sMaxThreads, out maxThreads))
                {
                    maxThreads = 3;
                }
                StringBuilder configString = new StringBuilder();
                configString.Append("{");
                configString.Append("\"zen_base_url\": \"").Append(smartNav.MetaWord(zenBaseURL)).Append("\",");
                configString.Append("\"zen_username\": \"").Append(smartNav.MetaWord(zenUserName)).Append("\",");
                configString.Append("\"zen_password\": \"").Append(smartNav.MetaWord(zenPassword)).Append("\",");
                configString.Append("\"adp_project_id\": \"").Append(smartNav.MetaWord(adpProjectID)).Append("\",");
                configString.Append("\"analyze_target\": \"").Append(smartNav.MetaWord(analyzeTarget)).Append("\",");
                configString.Append("\"json_options\": \"").Append(smartNav.MetaWord(jsonOptions)).Append("\",");
                configString.Append("\"login_target\": \"").Append(smartNav.MetaWord(loginTarget)).Append("\",");
                //String escapedOutputDirectoryPath = outputDirectoryPath;
                //if (escapedOutputDirectoryPath.Contains(@"\") && !escapedOutputDirectoryPath.Contains(@"\\"))
                //{
                //    escapedOutputDirectoryPath = escapedOutputDirectoryPath.Replace(@"\", @"\\");
                //}
                //configString.Append("\"output_directory_path\": \"").Append(smartNav.MetaWord(escapedOutputDirectoryPath)).Append("\",");
                configString.Append("\"output_options\": \"").Append(smartNav.MetaWord(outputOptions)).Append("\",");
                configString.Append("\"verify_token_target\": \"").Append(smartNav.MetaWord(verifyTokenTarget)).Append("\",");
                configString.Append("\"field_suffix\": \"").Append(smartNav.MetaWord(fieldSuffix)).Append("\",");
                configString.Append("\"timeout_in_minutes\": \"").Append(smartNav.MetaWord(timeoutInMinutes)).Append("\",");
                if ((UseAllPages != null) && UseAllPages.ToUpper().Trim().Equals("TRUE"))
                {
                    configString.Append("\"use_all_pages\": \"").Append(UseAllPages).Append("\",");
                }
                if ((smartNav.MetaWord(docClass) != null) && (smartNav.MetaWord(docClass).Trim().Length > 0))
                {
                    configString.Append("\"docClass\": \"").Append(smartNav.MetaWord(docClass)).Append("\",");
                }
                configString.Append("\"multiThreading\": \"").Append("true").Append("\"");
                configString.Append(" }");
                //RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "SendPageToADP: configString: " + configString.ToString());
                sendToADP(configString.ToString());
                utils.getRidOfUnneededADPFields(oPage, adpConfidenceAction);
            }
            catch (Exception e)
            {
                RRLog.WriteEx(OSADPConnector.LOGGING_ERROR, "error : " + e.Message);
                RRLog.WriteEx(OSADPConnector.LOGGING_ERROR, e.ToString());
                String className = this.GetType().Name;
                String methodName = System.Reflection.MethodInfo.GetCurrentMethod().Name;
                RRLog.WriteEx(OSADPConnector.LOGGING_ERROR, "Class " + className + ", method " + methodName);
                RRLog.WriteEx(OSADPConnector.LOGGING_ERROR, "Stack Trace: " + e.StackTrace);
            }
            return true;
        }

        public /*static*/ String uploadToADP(HttpClient httpClient, String url, ADPConfig adpConfig, ADPLoginResponse adpLR, int dpi, byte[] byteImage, String fileName, dclogXLib.IDCLog RRLog)
        {
            if (adpLR != null)
            {
                httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", adpLR.getAccessToken());
            }
            MultipartFormDataContent myMFDC = new MultipartFormDataContent("DCA-ADP-Upload----" + DateTime.Now.ToString(CultureInfo.InvariantCulture));
            String shortFN = Path.GetFileName(fileName);
            myMFDC.Add(new ByteArrayContent(byteImage), "file", shortFN);
            myMFDC.Add(new StringContent(adpConfig.output_options), "responseType");
            myMFDC.Add(new StringContent(adpConfig.json_options), "jsonOptions");
            if ((adpConfig.docClass != null) && (adpConfig.docClass.Trim().Length > 0))
            {
                myMFDC.Add(new StringContent(adpConfig.docClass), "docClass");
            }
            RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "RestServices: uploadToADP: Before PostAsync: Multipart Form Data Headers: " + myMFDC.Headers);
            RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "RestServices: uploadToADP: Before PostAsync: Multipart Form Data content (500 bytes): " + Environment.NewLine + myMFDC.ReadAsStringAsync().Result.Substring(0, 500));
            HttpResponseMessage response = httpClient.PostAsync(url, myMFDC).Result;

            /*
            processADP: 
            data: {
                "status": {
                    "code": 202, 
                    "messageId": "CIWCA50000", 
                    "message": "Success"
                }, 
                "result": [{
                    "status": { 
                        "code": 202, 
                        "messageId": "CIWCA11106", 
                        "message": "Content Analyzer request was created"
                    }, 
                    "data": {
                        "message": "json processing request was created successful", 
                        "fileNameIn": "TM000001.tif", 
                        "analyzerId": "44bf5d00-0aab-11ec-a241-9742b691a479", 
                        "type": ["json"]
                    }
                }]
            }
            */

            byte[] bytes = null;
            String analyzerId = null;

            if (response.IsSuccessStatusCode)
            {
                bytes = response.Content.ReadAsByteArrayAsync().Result;
                if (RRLog != null)
                {
                    RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "RestServices: uploadToADP: HTTP Code: " + response.StatusCode);
                    RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "RestServices: uploadToADP: Reason: " + response.ReasonPhrase);
                    RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "RestServices: uploadToADP: Content: " + response.Content.ReadAsStringAsync().Result);
                    RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "RestServices: uploadToADP: Headers: " + response.Headers.ToString());
                    if (response.RequestMessage != null)
                    {
                        RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "RestServices: uploadToADP: Request Message: " + response.RequestMessage.ToString());
                    }
                    RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "RestServices: uploadToADP: Response Message: " + response.ToString());
                }
                String fullResponse = Encoding.UTF8.GetString(bytes);
                String analyzerStart = "\"analyzerId\": \"";
                int analyzerStartIndex = fullResponse.IndexOf(analyzerStart);
                RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "RestServices: uploadToADP: index of '" + analyzerStart + "' is " + analyzerStartIndex);
                if (analyzerStartIndex < 0)
                {
                    analyzerStart = "\"analyzerId\":\"";
                    analyzerStartIndex = fullResponse.IndexOf(analyzerStart);
                    RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "RestServices: uploadToADP: removed space -- index of '" + analyzerStart + "' is " + analyzerStartIndex);
                }
                int start = fullResponse.IndexOf(analyzerStart) + analyzerStart.Length;
                int end = fullResponse.IndexOf("\"", start);
                analyzerId = fullResponse.Substring(start, end - start);
                RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "RestServices: uploadToADP: analyzerId: " + analyzerId);
            }
            else
            {
                // Some sort of error with first send, report the message in the log
                if (RRLog != null)
                {
                    RRLog.WriteEx(OSADPConnector.LOGGING_ERROR, "RestServices: uploadToADP: HTTP Code: " + response.StatusCode);
                    RRLog.WriteEx(OSADPConnector.LOGGING_ERROR, "RestServices: uploadToADP: Reason: " + response.ReasonPhrase);
                    RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "RestServices: uploadToADP: Content: " + response.Content.ReadAsStringAsync().Result);
                    RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "RestServices: uploadToADP: Headers: " + response.Headers.ToString());
                    if (response.RequestMessage != null)
                    {
                        RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "RestServices: uploadToADP: Request Message: " + response.RequestMessage.ToString());
                    }
                    RRLog.WriteEx(OSADPConnector.LOGGING_ERROR, "RestServices: uploadToADP: Response Message: " + response.ToString());
                }
                throw new ADPConnectorException(response.ReasonPhrase);

            }

            return analyzerId;

        }

        public /*static*/ void waitUntilDoneInADP(HttpClient httpClient, String url, ADPLoginResponse adpLR, dclogXLib.IDCLog RRLog, int timeoutInMinutes, out Boolean retry, out Boolean success)
        {
            success = false;
            retry = false;
            int secondsToWait = pollingIntervalInSeconds;
            int numLoops = ((timeoutInMinutes * 60) / secondsToWait) + 1;
            if (RRLog != null)
            {
                RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "RestServices: waitUntilDoneInADP: looping " + numLoops + " times, waiting " + secondsToWait + " seconds esach time");
            }
            Boolean done = false;
            for (int i = 0; i < numLoops && !done; i++)
            {
                //
                // Get this iteration's results
                //
                if (adpLR != null)
                {
                    httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", adpLR.getAccessToken());
                }

                HttpResponseMessage response = httpClient.GetAsync(url).Result;

                byte[] bytes = null;

                if (response.IsSuccessStatusCode)
                {
                    bytes = response.Content.ReadAsByteArrayAsync().Result;
                    if (RRLog != null)
                    {
                        RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "RestServices: waitUntilDoneInADP: HTTP Code: " + response.StatusCode);
                        RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "RestServices: waitUntilDoneInADP: Reason: " + response.ReasonPhrase);
                        RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "RestServices: waitUntilDoneInADP: Content: " + response.Content.ReadAsStringAsync().Result);
                        RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "RestServices: waitUntilDoneInADP: Headers: " + response.Headers.ToString());
                        if (response.RequestMessage != null)
                        {
                            RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "RestServices: waitUntilDoneInADP: Request Message: " + response.RequestMessage.ToString());
                        }
                        RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "RestServices: waitUntilDoneInADP: Response Message: " + response.ToString());
                    }
                    String fullResponse = Encoding.UTF8.GetString(bytes);
                    RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "RestServices: waitUntilDoneInADP: Full Response Message: " + fullResponse);
                    ADPProcessingStatus adpStatus = null;
                    try
                    {
                        adpStatus = JsonConvert.DeserializeObject<ADPProcessingStatus>(fullResponse);
                    }
                    catch (Exception e)
                    {
                        ADPProcessingStatus2 adpStatus2 = JsonConvert.DeserializeObject<ADPProcessingStatus2>(fullResponse);
                        adpStatus = adpStatus2.cloneToADPProcessingStatus();
                    }
                    if (adpStatus.status.code != 200)
                    {
                        if (RRLog != null)
                        {
                            RRLog.WriteEx(OSADPConnector.LOGGING_ERROR, "RestServices: waitUntilDoneInADP: adpStatus.status.code: " + adpStatus.status.code);
                            RRLog.WriteEx(OSADPConnector.LOGGING_ERROR, "RestServices: waitUntilDoneInADP: Reason: " + response.ReasonPhrase);
                            RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "RestServices: waitUntilDoneInADP: Content: " + response.Content.ReadAsStringAsync().Result);
                            RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "RestServices: waitUntilDoneInADP: Headers: " + response.Headers.ToString());
                            if (response.RequestMessage != null)
                            {
                                RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "RestServices: waitUntilDoneInADP: Request Message: " + response.RequestMessage.ToString());
                            }
                            RRLog.WriteEx(OSADPConnector.LOGGING_ERROR, "RestServices: waitUntilDoneInADP: Response Message: " + fullResponse);
                        }
                        throw new ADPConnectorException(response.ReasonPhrase);
                    }
                    foreach (ADPProcessingStatus.Result thisResult in adpStatus.result)
                    {
                        foreach (ADPProcessingStatus.StatusDetail sd in thisResult.data.statusDetails)
                        {
                            if (sd.status.ToLower().Trim().Equals("failed"))
                            {
                                if (RRLog != null)
                                {
                                    RRLog.WriteEx(OSADPConnector.LOGGING_ERROR, "RestServices: waitUntilDoneInADP: adpStatus.status.code: " + adpStatus.status.code);
                                    RRLog.WriteEx(OSADPConnector.LOGGING_ERROR, "RestServices: waitUntilDoneInADP: Reason: " + response.ReasonPhrase);
                                    RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "RestServices: waitUntilDoneInADP: Content: " + response.Content.ReadAsStringAsync().Result);
                                    RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "RestServices: waitUntilDoneInADP: Headers: " + response.Headers.ToString());
                                    if (response.RequestMessage != null)
                                    {
                                        RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "RestServices: waitUntilDoneInADP: Request Message: " + response.RequestMessage.ToString());
                                    }
                                    RRLog.WriteEx(OSADPConnector.LOGGING_ERROR, "RestServices: waitUntilDoneInADP: Response Message: " + fullResponse);
                                }
                                throw new ADPConnectorException(response.ReasonPhrase);
                            }
                            else if (sd.status.ToLower().Trim().Equals("completed"))
                            {
                                done = true;
                            }
                        }

                    }
                }
                else
                {
                    // Some sort of error with first send, report the message in the log
                    if (RRLog != null)
                    {
                        RRLog.WriteEx(OSADPConnector.LOGGING_ERROR, "RestServices: waitUntilDoneInADP: HTTP Code: " + response.StatusCode);
                        RRLog.WriteEx(OSADPConnector.LOGGING_ERROR, "RestServices: waitUntilDoneInADP: Reason: " + response.ReasonPhrase);
                        RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "RestServices: waitUntilDoneInADP: Content: " + response.Content.ReadAsStringAsync().Result);
                        RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "RestServices: waitUntilDoneInADP: Headers: " + response.Headers.ToString());
                        if (response.RequestMessage != null)
                        {
                            RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "RestServices: waitUntilDoneInADP: Request Message: " + response.RequestMessage.ToString());
                        }
                        RRLog.WriteEx(OSADPConnector.LOGGING_ERROR, "RestServices: waitUntilDoneInADP: Response Message: " + response.ToString());
                    }
                    // If we've timed out the connection, set up to retry the document
                    if (response.StatusCode.ToString().Trim().ToLower().Equals("unauthorized"))
                    {
                        retry = true;
                        done = true;
                    }
                    else
                    {
                        throw new ADPConnectorException(response.ReasonPhrase);
                    }

                }


                //
                // If not yet done, wait
                //
                if (!done)
                {
                    RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "RestServices: waitUntilDoneInADP: Waiting for " + secondsToWait + " seconds, loop " + i);
                    Thread.Sleep(secondsToWait * 1000);
                }
            }
            if (done && !retry)
            {
                success = true;
            }

            return;

        }

        public /*static*/ String getResultFromADP(HttpClient httpClient, String url, ADPLoginResponse adpLR, dclogXLib.IDCLog RRLog)
        {
            //
            // Get the analyzer results
            //
            if (adpLR != null)
            {
                httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", adpLR.getAccessToken());
            }

            HttpResponseMessage response = httpClient.GetAsync(url).Result;

            byte[] bytes = null;
            String fullResponse = null;

            if (response.IsSuccessStatusCode)
            {
                bytes = response.Content.ReadAsByteArrayAsync().Result;
                if (RRLog != null)
                {
                    RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "RestServices: getResultFromADP: HTTP Code: " + response.StatusCode);
                    RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "RestServices: getResultFromADP: Reason: " + response.ReasonPhrase);
                    RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "RestServices: getResultFromADP: Content: " + response.Content.ReadAsStringAsync().Result);
                    RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "RestServices: getResultFromADP: Headers: " + response.Headers.ToString());
                    if (response.RequestMessage != null)
                    {
                        RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "RestServices: getResultFromADP: Request Message: " + response.RequestMessage.ToString());
                    }
                    RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "RestServices: getResultFromADP: Response Message: " + response.ToString());
                }
                fullResponse = Encoding.UTF8.GetString(bytes);
                RRLog.WriteEx(OSADPConnector.LOGGING_ERROR, "RestServices: getResultFromADP: fullResponse: " + fullResponse);
            }
            else
            {
                // Some sort of error with first send, report the message in the log
                if (RRLog != null)
                {
                    RRLog.WriteEx(OSADPConnector.LOGGING_ERROR, "RestServices: getResultFromADP: HTTP Code: " + response.StatusCode);
                    RRLog.WriteEx(OSADPConnector.LOGGING_ERROR, "RestServices: getResultFromADP: Reason: " + response.ReasonPhrase);
                    RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "RestServices: getResultFromADP: Content: " + response.Content.ReadAsStringAsync().Result);
                    RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "RestServices: getResultFromADP: Headers: " + response.Headers.ToString());
                    if (response.RequestMessage != null)
                    {
                        RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "RestServices: getResultFromADP: Request Message: " + response.RequestMessage.ToString());
                    }
                    RRLog.WriteEx(OSADPConnector.LOGGING_ERROR, "RestServices: getResultFromADP: Response Message: " + response.ToString());
                }
                throw new ADPConnectorException(response.ReasonPhrase);

            }

            return fullResponse;

        }

        public /*static*/ void deleteFromADP(HttpClient httpClient, String url, ADPLoginResponse adpLR, dclogXLib.IDCLog RRLog)
        {
            //
            // Get the analyzer results
            //
            if (adpLR != null)
            {
                httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", adpLR.getAccessToken());
            }

            HttpResponseMessage response = httpClient.DeleteAsync(url).Result;

            byte[] bytes = null;
            String fullResponse = null;

            if (response.IsSuccessStatusCode)
            {
                bytes = response.Content.ReadAsByteArrayAsync().Result;
                if (RRLog != null)
                {
                    RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "RestServices: deleteFromADP: HTTP Code: " + response.StatusCode);
                    RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "RestServices: deleteFromADP: Reason: " + response.ReasonPhrase);
                    RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "RestServices: deleteFromADP: Content: " + response.Content.ReadAsStringAsync().Result);
                    RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "RestServices: deleteFromADP: Headers: " + response.Headers.ToString());
                    if (response.RequestMessage != null)
                    {
                        RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "RestServices: deleteFromADP: Request Message: " + response.RequestMessage.ToString());
                    }
                    RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "RestServices: deleteFromADP: Response Message: " + response.ToString());
                }
                fullResponse = Encoding.UTF8.GetString(bytes);
                RRLog.WriteEx(OSADPConnector.LOGGING_ERROR, "RestServices: deleteFromADP: fullResponse: " + fullResponse);
            }
            else
            {
                // Some sort of error with first send, report the message in the log
                if (RRLog != null)
                {
                    RRLog.WriteEx(OSADPConnector.LOGGING_ERROR, "RestServices: deleteFromADP: HTTP Code: " + response.StatusCode);
                    RRLog.WriteEx(OSADPConnector.LOGGING_ERROR, "RestServices: deleteFromADP: Reason: " + response.ReasonPhrase);
                    RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "RestServices: deleteFromADP: Content: " + response.Content.ReadAsStringAsync().Result);
                    RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "RestServices: deleteFromADP: Headers: " + response.Headers.ToString());
                    if (response.RequestMessage != null)
                    {
                        RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "RestServices: deleteFromADP: Request Message: " + response.RequestMessage.ToString());
                    }
                    RRLog.WriteEx(OSADPConnector.LOGGING_ERROR, "RestServices: deleteFromADP: Response Message: " + response.ToString());
                }
                throw new ADPConnectorException(response.ReasonPhrase);

            }

            return;

        }

        private void processADPJson(String jsonResults, /*Annotator annotator,*/ String adpFieldSuffix, TDCOLib.IDCO oPage, String pageID, int jsonPage)
        {
            // Get the ADP results
            ADPAnalyzerResults aar = new ADPAnalyzerResults(jsonResults, RRLog);
            String ocrText = aar.getOCRText();
            RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "processADP: ocrText for page " + jsonPage + ": " + ocrText);

            List<Tuple<String, String>> docClasses = aar.getDocumentClasses();

            RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "processADP: getting Blocks for page " + jsonPage);
            List<FieldObject> blocks = aar.getBlocks(jsonPage);
            int blockCount = 0;
            foreach (FieldObject block in blocks)
            {
                RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "processADP: block " + blockCount++ + " conf " + block.getConfidence() + " location (" + block.getLocation() + "): " + block.getAllLines());
                FieldObject[] lines = block.getChildren();
                int lineCount = 0;
                foreach (FieldObject line in lines)
                {
                    RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "  processADP: line " + lineCount++ + " conf " + line.getConfidence() + " location (" + line.getLocation() + "): " + line.getAllLines());
                    int wordCount = 0;
                    foreach (FieldObject child in line.getChildren())
                    {
                        RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "    processADP: word " + wordCount++ + " conf " + child.getConfidence() + " location (" + child.getLocation() + "): " + child.getAllLines());
                    }
                }
            }

            RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "processADP: getting Tables for page " + jsonPage);
            List<FieldObject> tables = aar.getTables(jsonPage);
            int tableCount = 0;
            foreach (FieldObject table in tables)
            {
                RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "processADP: table " + tableCount++ + " conf " + table.getConfidence() + " location (" + table.getLocation() + "): " + table.getAllLines());
                FieldObject[] rows = table.getChildren();
                int rowCount = 0;
                foreach (FieldObject row in rows)
                {
                    RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "  processADP: row " + rowCount++ + " conf " + row.getConfidence() + " location (" + row.getLocation() + "): " + row.getAllLines());
                    int cellCount = 0;
                    foreach (FieldObject cell in row.getChildren())
                    {
                        RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "    processADP: cell " + cellCount++ + " conf " + cell.getConfidence() + " location (" + cell.getLocation() + "): " + cell.getAllLines());
                        int wordCount = 0;
                        foreach (FieldObject child in cell.getChildren())
                        {
                            RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "    processADP: word " + wordCount++ + " conf " + child.getConfidence() + " location (" + child.getLocation() + "): " + child.getAllLines());
                        }
                    }
                }
            }

            ADPPageDimensions adpPageDim = aar.getPageDimensions();
            //RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "processADP: Page dimensions: ");
            //RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "            Width: " + adpPageDim.PageWidth);
            //RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "            Height: " + adpPageDim.PageHeight);
            //RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "            OCR Confidence: " + adpPageDim.PageOCRCOnfidence);
            //RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "            dpiX: " + adpPageDim.dpiX);
            //RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "            dpiY: " + adpPageDim.dpiY);

            List<ADPRanking> adpRankings = aar.getRankings();

            RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "processADP: getting Key-Value Pairs for page " + jsonPage);
            List<ADPKeyValuePair> kvpList = aar.getKeyValuePairs(adpRankings, adpFieldsAction, jsonPage);
            //RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "processADP: Key Value Pairs: ");
            //int kvpNum = 0;
            //foreach (ADPKeyValuePair kvp in kvpList)
            //{
            //    RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "    KVP " + kvpNum++);
            //    RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "       Entity Type: " + kvp.entityType);
            //    RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "       Entity Confidence: " + kvp.keyClassConfidence);
            //    RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "       Key: " + kvp.key);
            //    RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "       Key X: " + kvp.keyX);
            //    RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "       Key Y: " + kvp.keyY);
            //    RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "       Key Width: " + kvp.keyWidth);
            //    RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "       Key Height: " + kvp.keyHeight);
            //    RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "       Value: " + kvp.value);
            //    RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "       Value X: " + kvp.valueX);
            //    RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "       Value Y: " + kvp.valueY);
            //    RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "       Value Width: " + kvp.valueWidth);
            //    RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "       Value Height: " + kvp.valueWidth);
            //    RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "       Original Key: " + kvp.originalKey);
            //    RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "       Original Value: " + kvp.originalValue);
            //}

            RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "processADP: getting Table Key-Value Pairs for page " + jsonPage);
            List<ADPKeyValuePair> tableKVPList = aar.getTableKeyValuePairs(jsonPage);
            //RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "processADP: Table Key Value Pairs: ");
            //int tkvpNum = 0;
            //foreach (ADPKeyValuePair kvp in tableKVPList)
            //{
            //    RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "    Table " + tkvpNum++);
            //    RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "       Entity Type: " + kvp.entityType);
            //    RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "       Key: " + kvp.key);
            //    RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "       Key X: " + kvp.keyX);
            //    RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "       Key Y: " + kvp.keyY);
            //    RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "       Key Width: " + kvp.keyWidth);
            //    RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "       Key Height: " + kvp.keyHeight);
            //    RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "       Value: " + kvp.value);
            //    RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "       Value X: " + kvp.valueX);
            //    RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "       Value Y: " + kvp.valueY);
            //    RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "       Value Width: " + kvp.valueWidth);
            //    RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "       Value Height: " + kvp.valueWidth);
            //    RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "       Original Key: " + kvp.originalKey);
            //    RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "       Original Value: " + kvp.originalValue);
            //    int liNum = 0;
            //    foreach (ADPKeyValuePair lineItem in kvp.nested)
            //    {
            //        RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "       Line Item " + liNum++);
            //        int cellNum = 0;
            //        foreach (ADPKeyValuePair cell in lineItem.nested)
            //        {
            //            RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "          Cell " + cellNum++);
            //            RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "             Entity Type: " + cell.entityType);
            //            RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "             Key: " + cell.key);
            //            RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "             Key X: " + cell.keyX);
            //            RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "             Key Y: " + cell.keyY);
            //            RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "             Key Width: " + cell.keyWidth);
            //            RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "             Key Height: " + cell.keyHeight);
            //            RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "             Value: " + cell.value);
            //            RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "             Value X: " + cell.valueX);
            //            RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "             Value Y: " + cell.valueY);
            //            RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "             Value Width: " + cell.valueWidth);
            //            RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "             Value Height: " + cell.valueWidth);
            //            RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "             Original Key: " + cell.originalKey);
            //            RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "             Original Value: " + cell.originalValue);
            //        }
            //    }
            //}
            //RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "after parsing adp json");

            buildLayoutXMLFromBlocksAndTables(blocks, tables, pageID, adpPageDim.dpiX, adpPageDim.dpiY, adpPageDim.PageWidth, adpPageDim.PageHeight);

            RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "after reading building layout xml");

            String classificationVarName = ADPConfig.ADPDocType;
            addADPFields(oPage, aar, classificationVarName, docClasses, kvpList, tableKVPList, adpFieldSuffix);

            RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "after adding normal fields");

            // Add the layout variable to the page, Abbyy Action was doing this behind the scenes
            oPage.set_Variable("layout", pageID + "_layout.xml");

            RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "after adding table fields");


        }

        private IDictionary<Int32, Int32> findFontSizesFromWordsOrLines(IList<FieldObject> fields, Boolean fromADP = false)
        {
            IDictionary<Int32, Int32> fontSizes = new Dictionary<Int32, Int32>();

            for (int i = 0; i < fields.Count; i++)
            {
                FieldObject field = fields[i];

                if (field.getFieldType() == FieldObject.FIELD_TYPE_WORD)
                {
                    int fontSize = field.getFontSize();
                    if (fromADP && (fontSize % 100 == 0))
                    {
                        fontSize = fontSize / 100;
                    }
                    if (!fontSizes.ContainsKey(fontSize)) // Only add new sizes
                    {
                        fontSizes.Add(fontSize, fontSize);
                    }
                }
                else if (field.getFieldType() == FieldObject.FIELD_TYPE_WORDGROUP)
                {
                    FieldObject[] words = field.getChildren(FieldObject.FIELD_TYPE_WORD);

                    for (int a = 0; a < words.Length; a++)
                    {
                        FieldObject f = words[a];
                        int fontSize = field.getFontSize();
                        if (fromADP && (fontSize % 100 == 0))
                        {
                            fontSize = fontSize / 100;
                        }
                        if (!fontSizes.ContainsKey(fontSize)) // Only add new sizes
                        {
                            fontSizes.Add(fontSize, fontSize);
                        }

                    }
                }
                else
                {
                    FieldObject[] children = field.getChildren();
                    IDictionary<Int32, Int32> newFontSizes = findFontSizesFromWordsOrLines(children, fromADP);
                    foreach (var pair in newFontSizes)
                    {
                        if (!fontSizes.ContainsKey(pair.Key))
                        {
                            fontSizes.Add(pair.Key, pair.Value);
                        }
                    }
                }
            }

            return fontSizes;
        }

        private void buildLayoutXMLFromBlocksAndTables(List<FieldObject> blocks, List<FieldObject> tables, String pageID, int xDpi, int yDpi, int width, int height)
        {
            List<FieldObject> wordGroups = new List<FieldObject>();
            foreach (FieldObject block in blocks)
            {
                wordGroups.Add(block);
            }
            foreach (FieldObject table in tables)
            {
                // tables contain rows
                FieldObject[] rows = table.getChildren();
                // rows contain cells
                foreach (FieldObject row in rows)
                {
                    FieldObject[] cells = row.getChildren();
                    // cells contain lines, which contain words
                    foreach (FieldObject cell in cells)
                    {
                        FieldObject[] lines = cell.getChildren();
                        foreach (FieldObject line in lines)
                        {
                            wordGroups.Add(line);
                        }
                    }
                }
            }
            IDictionary<Int32, Int32> fontSizes = this.findFontSizesFromWordsOrLines(wordGroups, true);

            List<FieldObject> mergedBlocksAndTables = sortBlocksAndTables(blocks, tables);

            writeLayoutXML(mergedBlocksAndTables.ToArray(), fontSizes, pageID, xDpi, yDpi, width, height);

        }

        List<FieldObject> sortBlocksAndTables(List<FieldObject> blocks, List<FieldObject> tables)
        {
            // Build Dictionary for WordGroups
            List<FieldObject> sorted = new List<FieldObject>();
            IDictionary<PageLocation, FieldObject> foDictionary = new Dictionary<PageLocation, FieldObject>();

            foreach (FieldObject block in blocks)
            {
                foDictionary.Add(block.getLocation(), block);
            }
            foreach (FieldObject table in tables)
            {
                foDictionary.Add(table.getLocation(), table);
            }

            while (true)
            {
                FieldObject[] fos = allFieldsToArray(foDictionary);
                PageLocation location = findClosestFieldToTheLeftAndTop(fos);

                if (location == null) // Break out of the loop if nothing is left
                {
                    break;
                }

                FieldObject firstFO = foDictionary[location];
                foDictionary.Remove(location); // Remove the field from the dictionary
                sorted.Add(firstFO);

            }

            return sorted;
        }

        private FieldObject[] allFieldsToArray(IDictionary<PageLocation, FieldObject> fields)
        {
            IList<FieldObject> fos = new List<FieldObject>();
            foreach (PageLocation location in fields.Keys)
            {
                fos.Add(fields[location]);
            }

            return fos.ToArray();
        }

        private PageLocation findClosestFieldToTheLeftAndTop(FieldObject[] fields)
        {
            int closestLeft = 100000;
            int closestTop = 100000;
            PageLocation location = null;

            for (int i = 0; i < fields.Length; i++)
            {
                if (fields[i].getLocation().getTop() < closestTop)
                {
                    if (fields[i].getLocation().getLeft() < closestLeft)
                    {
                        closestLeft = fields[i].getLocation().getLeft();
                        closestTop = fields[i].getLocation().getTop();
                        location = fields[i].getLocation();
                    }
                }
            }

            return location;
        }

        private void writeLayoutXML(FieldObject[] blocks, IDictionary<Int32, Int32> fontSizes, String pageID, int xDpi, int yDpi, int width, int height)
        {

            IDictionary<Int32, Int32> newFontSizes = new Dictionary<Int32, Int32>();

            int cntr = 0; // Set IDs for the fontSizes
            foreach (Int32 fontSize in fontSizes.Keys)
            {
                newFontSizes.Add(fontSize, cntr);
                cntr++;
            }

            StringBuilder xml = new StringBuilder();

            xml.Append("<?xml version=\"1.0\" encoding=\"utf-16\"?>\r\n");
            xml.Append("<Page xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xmlns:xsd=\"http://www.w3.org/2001/XMLSchema\" pos=\"0,0," + width + "," + height + "\" lang=\"English\" id=\"" + pageID + "\" printArea=\"0,0," + width + "," + height + "\" xdpi=\"" + xDpi + "\" ydpi=\"" + yDpi + "\" >\r\n");

            RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "writeLayoutXML: xml length before addFieldObjectsToSB: " + xml.Length);

            xml = addFieldObjectsToSB(xml, blocks, "", fontSizes, newFontSizes);

            RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "writeLayoutXML: xml length after addFieldObjectsToSB: " + xml.Length);

            // Styles
            foreach (Int32 fontSize in newFontSizes.Keys)
            {
                xml.Append("  <Style id=\"" + newFontSizes[fontSize] + "\" v=\"color: 000000; font-name: arial; font-family: ft_sansserif; font-size: " + fontSize + "pt; \" />\r\n");
            }

            xml.Append("</Page>\r\n");

            File.WriteAllText(BatchPilot.BatchDir + "\\" + pageID + "_layout.xml", xml.ToString(), Encoding.Unicode);

        }

        private StringBuilder addFieldObjectsToSB(StringBuilder xmlIn, FieldObject[] FOs, String pad, IDictionary<int, int> fontSizes, IDictionary<int, int> newFontSizes)
        {
            StringBuilder xml = new StringBuilder(xmlIn.ToString());
            String newPad = pad + "  ";
            for (int i = 0; i < FOs.Length; i++)
            {
                FieldObject thisFO = FOs[i];
                if (thisFO.getFieldType() == FieldObject.FIELD_TYPE_BLOCK)
                {
                    xml = addBlockToSB(xml, thisFO, newPad, fontSizes, newFontSizes);
                }
                else if (thisFO.getFieldType() == FieldObject.FIELD_TYPE_WORDGROUPBLOCK)
                {
                    xml = addBlockToSB(xml, thisFO, newPad, fontSizes, newFontSizes);
                }
                else if (thisFO.getFieldType() == FieldObject.FIELD_TYPE_CELL)
                {
                    xml = addCellToSB(xml, thisFO, newPad, fontSizes, newFontSizes);
                }
                else if (thisFO.getFieldType() == FieldObject.FIELD_TYPE_LINE)
                {
                    xml = addLineToSB(xml, thisFO, newPad, fontSizes, newFontSizes);
                }
                else if (thisFO.getFieldType() == FieldObject.FIELD_TYPE_WORDGROUP)
                {
                    xml = addLineToSB(xml, thisFO, newPad, fontSizes, newFontSizes);
                }
                else if (thisFO.getFieldType() == FieldObject.FIELD_TYPE_WORDGROUPLINE)
                {
                    xml = addLineToSB(xml, thisFO, newPad, fontSizes, newFontSizes);
                }
                else if (thisFO.getFieldType() == FieldObject.FIELD_TYPE_PARAGRAPH)
                {
                    xml = addParaToSB(xml, thisFO, newPad, fontSizes, newFontSizes);
                }
                else if (thisFO.getFieldType() == FieldObject.FIELD_TYPE_ROW)
                {
                    xml = addRowToSB(xml, thisFO, newPad, fontSizes, newFontSizes);
                }
                else if (thisFO.getFieldType() == FieldObject.FIELD_TYPE_TABLE)
                {
                    xml = addTableToSB(xml, thisFO, newPad, fontSizes, newFontSizes);
                }
                else if (thisFO.getFieldType() == FieldObject.FIELD_TYPE_WGTABLE)
                {
                    xml = addTableToSB(xml, thisFO, newPad, fontSizes, newFontSizes);
                }
                else if (thisFO.getFieldType() == FieldObject.FIELD_TYPE_WORD)
                {
                    Boolean moreWords = true;
                    if (i + 1 == FOs.Length)
                    {
                        moreWords = false;
                    }
                    xml = addWordToSB(xml, thisFO, newPad, fontSizes, newFontSizes, moreWords);
                }
                else if (thisFO.getFieldType() == FieldObject.FIELD_TYPE_ROWGROUP)
                {
                    //addRowGroupToSB(xml, thisFO, newPad, fontSizes, newFontSizes);
                }
            }
            return xml;
        }

        private StringBuilder addBlockToSB(StringBuilder xmlIn, FieldObject block, String pad, IDictionary<int, int> fontSizes, IDictionary<int, int> newFontSizes)
        {
            StringBuilder xml = new StringBuilder(xmlIn.ToString());

            xml.Append(pad).Append("<Block pos=\"" + block.getLocation().getPosition() + "\">\r\n");

            xml = addFieldObjectsToSB(xml, block.getChildren(), pad + "  ", fontSizes, newFontSizes);

            xml.Append(pad).Append("</Block>\r\n");

            return xml;
        }

        private StringBuilder addLineToSB(StringBuilder xmlIn, FieldObject line, String pad, IDictionary<int, int> fontSizes, IDictionary<int, int> newFontSizes)
        {
            StringBuilder xml = new StringBuilder(xmlIn.ToString());

            xml.Append(pad).Append("<L pos=\"" + line.getLocation().getPosition() + "\" s=\"" + getStyleIdForFontSize(getLineFontSize(line, fontSizes), newFontSizes) + "\">\r\n");

            xml = addFieldObjectsToSB(xml, line.getChildren(), pad + "  ", fontSizes, newFontSizes);

            xml.Append(pad).Append("</L>\r\n");

            return xml;
        }

        private StringBuilder addParaToSB(StringBuilder xmlIn, FieldObject para, String pad, IDictionary<int, int> fontSizes, IDictionary<int, int> newFontSizes)
        {
            StringBuilder xml = new StringBuilder(xmlIn.ToString());

            xml.Append(pad).Append("<Para pos=\"" + para.getLocation().getPosition() + "\" s=\"" + getStyleIdForFontSize(getLineFontSize(para, fontSizes), newFontSizes) + "\">\r\n");

            xml = addFieldObjectsToSB(xml, para.getChildren(), pad + "  ", fontSizes, newFontSizes);

            xml.Append(pad).Append("</Para>\r\n");

            return xml;
        }

        private StringBuilder addWordToSB(StringBuilder xmlIn, FieldObject word, String pad, IDictionary<int, int> fontSizes, IDictionary<int, int> newFontSizes, Boolean moreWords)
        {
            StringBuilder xml = new StringBuilder(xmlIn.ToString());

            xml.Append(pad).Append("<W pos=\"" + word.getLocation().getPosition() + "\" v=\"" + cleanAttributeText(word.getLines()[0])
                                    + "\" s=\"" + getStyleIdForFontSize(word.getFontSize(), newFontSizes) + "\" cn=\"" + getWordScore(word) + "\">\r\n");

            for (int b = 0; b < word.getLines()[0].Length; b++)
            {
                xml.Append(pad).Append("  ").Append("<C pos=\"" + findCharacterPosition(word.getLocation(), word.getLines()[0].Length, b)
                                                    + "\" v=\"" + cleanAttributeText(((word.getLines()[0])[b]) + "") + "\" s=\""
                                                    + getStyleIdForFontSize(word.getFontSize(), newFontSizes) + "\" cn=\""
                                                    + Convert.ToInt32(Math.Round(Convert.ToDouble(word.getConfidence() / 10))) + "\" />\r\n");
            }

            xml.Append(pad).Append("</W>\r\n");

            if (moreWords)
            {
                xml.Append(pad).Append("<S />\r\n");
            }

            return xml;
        }

        int cellNum = 0;
        int rowNum = 0;

        private StringBuilder addTableToSB(StringBuilder xmlIn, FieldObject table, String pad, IDictionary<int, int> fontSizes, IDictionary<int, int> newFontSizes)
        {
            StringBuilder xml = new StringBuilder(xmlIn.ToString());

            // table --> rows --> cells --> lines --> words
            FieldObject[] rows = table.getChildren();
            int numRows = rows.Length;
            int numCols = 0;
            if ((rows != null) && (rows.Length > 0))
            {
                FieldObject[] cols = rows[0].getChildren();
                numCols = cols.Length;
            }

            xml.Append(pad).Append("<Table columns=\"" + numCols + "\" pos=\"" + table.getLocation().getPosition() + "\" rows=\"" + numRows + "\">\r\n");
            cellNum = 0;
            rowNum = 0;

            xml = addFieldObjectsToSB(xml, table.getChildren(), pad + "  ", fontSizes, newFontSizes);

            xml.Append(pad).Append("</Table>\r\n");

            return xml;
        }

        private StringBuilder addRowToSB(StringBuilder xmlIn, FieldObject row, String pad, IDictionary<int, int> fontSizes, IDictionary<int, int> newFontSizes)
        {
            cellNum = 0;
            StringBuilder xml = new StringBuilder(xmlIn.ToString());

            xml.Append(pad).Append("<Row pos=\"" + row.getLocation().getPosition() + "\">\r\n");

            xml = addFieldObjectsToSB(xml, row.getChildren(), pad + "  ", fontSizes, newFontSizes);

            xml.Append(pad).Append("</Row>\r\n");
            rowNum++;

            return xml;
        }

        private StringBuilder addCellToSB(StringBuilder xmlIn, FieldObject cell, String pad, IDictionary<int, int> fontSizes, IDictionary<int, int> newFontSizes)
        {
            StringBuilder xml = new StringBuilder(xmlIn.ToString());

            xml.Append(pad).Append("<Cell col=\"" + cellNum + "\" pos=\"" + cell.getLocation().getPosition() + "\" row=\"" + rowNum + "\" columnSpan=\"1\">\r\n");

            xml = addFieldObjectsToSB(xml, cell.getChildren(), pad + "  ", fontSizes, newFontSizes);

            xml.Append(pad).Append("</Cell>\r\n");
            cellNum++;

            return xml;
        }

        private int getStyleIdForFontSize(int fontSize, IDictionary<Int32, Int32> fontSizes)
        {
            try
            {
                if (fontSizes.ContainsKey(fontSize))
                {
                    return fontSizes[fontSize];
                }
                int newFontSize = fontSize;
                if (fontSize % 100 == 0)
                {
                    newFontSize = fontSize / 100;
                }
                if (fontSizes.ContainsKey(newFontSize))
                {
                    return fontSizes[newFontSize];
                }
                RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "getStyleIdForFontSize: error: font size dictionary does not contain desired fontsize");
                RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "getStyleIdForFontSize: looking for font size " + fontSize);
                RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "getStyleIdForFontSize: font sizes in dictionary:");
                foreach (int thisFontSize in fontSizes.Keys)
                {
                    RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "getStyleIdForFontSize: " + thisFontSize);
                }
                return 0;
            }
            catch (Exception e)
            {
                RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "getStyleIdForFontSize: error: " + e.Message);
                RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "getStyleIdForFontSize: looking for font size " + fontSize);
                RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "getStyleIdForFontSize: font sizes in dictionary:");
                foreach (int thisFontSize in fontSizes.Keys)
                {
                    RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "getStyleIdForFontSize: " + thisFontSize);
                }
                throw e;
            }
        }

        private int getLineFontSize(FieldObject wordGroup, IDictionary<Int32, Int32> fontSizes)
        {
            FieldObject[] words = wordGroup.getChildren(FieldObject.FIELD_TYPE_WORD);
            int total = 0;

            for (int i = 0; i < words.Length; i++)
            {
                total += words[i].getFontSize();
            }

            if (words.Length == 0)
            {
                return 0;
            }

            Int32 fsize = total / words.Length;

            return findClosestSize(fsize, fontSizes);

        }

        private int findClosestSize(int size, IDictionary<Int32, Int32> fontSizes)
        {
            if (fontSizes.ContainsKey(size))
            {
                return size;
            }

            List<Int32> keyList = new List<Int32>(fontSizes.Keys);

            int closest = keyList.Aggregate((x, y) => Math.Abs(x - size) < Math.Abs(y - size) ? x : y);

            return closest;

        }

        private String cleanAttributeText(String value)
        {
            value = value.Replace("&", "&amp;")
                         .Replace("\"", "&quot;")
                         .Replace("'", "&apos;")
                         .Replace("<", "&lt;")
                         .Replace(">", "&gt;");

            return value;
        }


        private String findCharacterPosition(PageLocation location, int totalLength, int position)
        {
            int left = location.getLeft(); // 50
            int right = location.getRight(); // 60

            int distanceBetween = right - left; // 10
            int averageDistance = distanceBetween / totalLength; // Word 2.5

            int newLeft = left + (position * averageDistance);
            int newRight = newLeft + averageDistance;

            // ltrb
            return newLeft + "," + location.getTop() + "," + newRight + "," + location.getBottom();
        }

        private String getWordScore(FieldObject word)
        {
            String returnScore = "";

            for (int i = 0; i < word.getLines()[0].Length; i++)
            {
                Int32 newScore = Convert.ToInt32(Math.Round(Convert.ToDouble(word.getConfidence() / 10)));
                newScore = newScore - 1; // a value of 10 is possible for the characters, but for the word only a 9 is possible
                if (newScore > 9)
                {
                    newScore = 9;
                }

                if (newScore < 0)
                {
                    newScore = 0;
                }


                returnScore += (newScore + "");
            }

            return returnScore;
        }

        private void addADPFields(TDCOLib.IDCO oPage, ADPAnalyzerResults aar, String classVar, List<Tuple<String, String>> docClasses, List<ADPKeyValuePair> normalKVPs, List<ADPKeyValuePair> tableKVPs, String adpFieldSuffix)
        {
            RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "RestServicesAction addADPFields: start");

            // Add classification info to page as variables
            if ((docClasses != null) && (docClasses.Count > 0))
            {
                Tuple<String, String> firstTuple = docClasses[0];
                oPage.set_Variable(classVar, firstTuple.Item1);
                oPage.set_Variable(classVar + "Confidence", firstTuple.Item2);
                for (int i = 0; i < docClasses.Count; i++)
                {
                    String classVarExtended = classVar + "_" + i;
                    Tuple<String, String> thisTuple = docClasses[i];
                    oPage.set_Variable(classVarExtended, thisTuple.Item1);
                    oPage.set_Variable(classVarExtended + "Confidence", thisTuple.Item2);
                }
            }

            // Add normal fields in the following order:
            //   Key Class Confidence: High, Medium, Low
            //      entityName
            //         Key Position
            // Start from the top left, looking for high Key Class Confidence
            // When you find one, add it, then look for all others with that Key Class
            RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "RestServicesAction addADPFields: normal KVPs before sorting");
            aar.LogKVPList(normalKVPs, "");
            normalKVPs.Sort(); // sort by confidence and position
            RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "RestServicesAction addADPFields: normal KVPs after sorting");
            aar.LogKVPList(normalKVPs, "");

            String[] confidences = { "high", "medium", "low" };
            for (int confIndex = 0; confIndex < confidences.Length; confIndex++)
            {
                String confidence = confidences[confIndex];
                RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "RestServicesAction addADPFields: processing KeyClassConfidence=" + confidence);
                for (int i = 0; i < normalKVPs.Count; i++)
                {
                    ADPKeyValuePair thisKVP = normalKVPs[i];
                    RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "RestServicesAction addADPFields: processing KVP " + i + ", key class " + thisKVP.keyClass);
                    if (!thisKVP.added)
                    {
                        // Find the number of ADP entries matching this ADP key class
                        int numForThisKeyClass = 0;
                        for (int innerIndex = 0; innerIndex < normalKVPs.Count; innerIndex++)
                        {
                            ADPKeyValuePair innerKVP = normalKVPs[innerIndex];
                            if (thisKVP.keyClass.Trim().Equals(innerKVP.keyClass.Trim()))
                            {
                                numForThisKeyClass++;
                            }
                        }
                        ADPKeyValuePair kvp = normalKVPs[i];
                        if (kvp.keyClassConfidence != null)
                        {
                            if (!kvp.added && kvp.keyClassConfidence.Trim().ToLower().Equals(confidence))
                            {
                                setFieldInDCO(kvp, oPage, "", numForThisKeyClass, adpFieldSuffix);
                                kvp.added = true;
                                if (kvp.keyClass != null)
                                {
                                    for (int innerConfIndex = 0; innerConfIndex < confidences.Length; innerConfIndex++)
                                    {
                                        String innerConf = confidences[innerConfIndex];
                                        for (int j = i + 1; j < normalKVPs.Count; j++)
                                        {
                                            if (!normalKVPs[j].added)
                                            {
                                                ADPKeyValuePair otherKVP = normalKVPs[j];
                                                if (!otherKVP.added && (otherKVP.keyClassConfidence.Trim().ToLower().Equals(innerConf)) &&
                                                    (otherKVP.keyClass != null) && kvp.keyClass.Trim().ToLower().Equals(otherKVP.keyClass.Trim().ToLower()))
                                                {
                                                    setFieldInDCO(otherKVP, oPage, "", numForThisKeyClass, adpFieldSuffix);
                                                    otherKVP.added = true;
                                                    RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "RestServicesAction addADPFields: added other KVP item " + j + "key class " + otherKVP.keyClass);
                                                }
                                            }
                                        }
                                    }
                                }
                                else
                                {
                                    RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "RestServicesAction addADPFields: key class was null");
                                }
                            }
                            else
                            {
                                RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "RestServicesAction addADPFields: item already added or key class confidence (" + kvp.keyClassConfidence
                                    + ") doesn't match what we're looking for (" + confidence + ")");
                            }
                        }
                        else
                        {
                            RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "RestServicesAction addADPFields: item i=" + i + " had no KeyClassConfidence");
                        }
                    }
                    else
                    {
                        RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "RestServicesAction addADPFields: item had already been added");
                    }
                }
            }

            // Add table fields
            if ((tableKVPs != null) && (tableKVPs.Count > 0))
            {
                foreach (ADPKeyValuePair table in tableKVPs)
                {
                    // Find the number of ADP entries matching this ADP key class
                    int numForThisKeyClass = 0;
                    ADPKeyValuePair thisKVP = table;
                    for (int innerIndex = 0; innerIndex < normalKVPs.Count; innerIndex++)
                    {
                        ADPKeyValuePair innerKVP = normalKVPs[innerIndex];
                        if (thisKVP.keyClass.Trim().Equals(innerKVP.keyClass.Trim()))
                        {
                            numForThisKeyClass++;
                        }
                    }
                    TDCOLib.IDCO oTable = setFieldInDCO(table, oPage, "", numForThisKeyClass, adpFieldSuffix);
                    if (oTable != null)
                    {
                        List<ADPKeyValuePair> rows = table.nested;
                        if ((rows != null) && (rows.Count > 0))
                        {
                            int rownum = 0;
                            foreach (ADPKeyValuePair row in rows)
                            {
                                TDCOLib.IDCO oRow = setFieldInDCO(row, oTable, rownum.ToString(), 0, adpFieldSuffix, "Lineitem");
                                rownum++;
                                if (oRow != null)
                                {
                                    List<ADPKeyValuePair> cells = row.nested;
                                    if ((cells != null) && (cells.Count > 0))
                                    {
                                        foreach (ADPKeyValuePair cell in cells)
                                        {
                                            TDCOLib.IDCO cellField = setFieldInDCO(cell, oRow, "", 0, adpFieldSuffix);
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
            RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "RestServicesAction addADPFields: end");
        }

        private TDCOLib.IDCO setFieldInDCO(ADPKeyValuePair kvp, TDCOLib.IDCO dco, String suffix, int numForThisKeyClass, String adpFieldSuffix, String nameOverride = null)
        {
            RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "RestServicesAction setFieldInDCO: have " + numForThisKeyClass + " fields for key class " + kvp.keyClass + ".");
            if (dco != null)
            {
                //Entity entity = new Entity();
                String nameToUse = kvp.keyClass;
                if (nameOverride != null)
                {
                    nameToUse = nameOverride;
                }
                if ((nameToUse == null) || (nameToUse.Trim().Length == 0))
                {
                    nameToUse = kvp.key;
                }
                String typeName = nameToUse;
                // Changed this -- always add the "_ADP_nn" to ADP
                int nameIndex = 0;
                Boolean found = true;
                while (found)
                {
                    String nameToCheck = nameToUse + adpFieldSuffix;
                    if (numForThisKeyClass > 1)
                    {
                        nameToCheck = nameToCheck + "_" + nameIndex;
                    }
                    else if (numForThisKeyClass == 1)
                    {
                        // Exactly 1 entry for this key class, so should only be a single field on the page -- this one
                        found = false;
                    }
                    if (dco.FindChildIndex(nameToCheck) < 0)
                    {
                        found = false;
                        nameToUse = nameToCheck;
                    }
                    else
                    {
                        nameIndex++;
                        if (nameIndex > 100)
                        {
                            found = false;
                            RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "setFieldInDCO: already added more than 100 iterations of " + nameToUse + ", just checked " + nameToCheck + ", skipping this key value pair from ADP");
                        }
                    }
                }
                nameToUse = nameToUse + suffix;
                PageLocation keyLocation = new PageLocation(kvp.keyX, kvp.keyY, kvp.keyX + kvp.keyWidth, kvp.keyY + kvp.keyHeight);
                PageLocation valueLocation = new PageLocation(kvp.valueX, kvp.valueY, kvp.valueX + kvp.valueWidth, kvp.valueY + kvp.valueHeight);
                utils.updateKeyMatchEntityAndProperties(dco, nameToUse, nameToUse, "ADP", kvp.key, keyLocation, kvp.confidence, null, -1, typeName);
                utils.updateEntityAndProperties(dco, nameToUse, nameToUse, "ADP", kvp.value, valueLocation, kvp.confidence, null, null, null, -1, typeName);
                TDCOLib.IDCO newField = dco.FindChild(nameToUse);
                newField.set_Variable(ADPKeyValuePair.ADPKeyClass, kvp.keyClass);
                newField.set_Variable(ADPKeyValuePair.ADPKeyClassConfidence, kvp.keyClassConfidence);
                newField.set_Variable(ADPKeyValuePair.ADPSensitivity, kvp.sensitivity.ToString());
                if (kvp.haveLineItem)
                {
                    newField.set_Variable(ADPKeyValuePair.ADPLineItemID, kvp.LineItemID.ToString());
                    newField.set_Variable(ADPKeyValuePair.ADPSeqLineItemID, kvp.SeqLineItemID.ToString());
                }
                return newField;
            }
            return null;
        }

        // Document level action
        // Assumes the first page-level object within the document is a multi-page object that has been sent to ADP for processing
        // Assumes the following page-level objects in the document are the individual pages from the multi-page object
        // Reads the ADP results from the first page-level object and applies them to the following individual pages
        public bool SpreadTheADPResults(String fieldsAction, String confidenceAction)
        {
            DateTime compileTime = new DateTime(Builtin.CompileTime, DateTimeKind.Utc); // Ignore the supposed error -- this is resolved at compile time
            RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "***Start SpreadTheADPResults - compile time " + compileTime);
            if (smartNav == null)
            {
                smartNav = new dcSmart.SmartNav(this);
            }
            Boolean status = false;

            int adpFieldsAction = convertFieldsString(smartNav.MetaWord(fieldsAction));
            RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "SpreadTheADPResults: after conveertFieldsString");
            int adpConfidenceAction = convertConfidenceString(smartNav.MetaWord(confidenceAction));
            RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "SpreadTheADPResults: after convertConfidenceString");

            this.adpFieldsAction = adpFieldsAction;

            try
            {
                if (smartNav == null)
                {
                    smartNav = new dcSmart.SmartNav(this);
                }
                else
                {
                    standalone = true;
                }
                this.utils = new Utils(RRLog);
                TDCOLib.IDCO oDoc = this.CurrentDCO;
                if (oDoc.ObjectType() != Level.Document)
                {
                    throw new Exception("SpreadTheADPResults must be at a document level");
                }
                if (oDoc.NumOfChildren() < 1)
                {
                    throw new Exception("SpreadTheADPResults must be invoked for a non-empty document");
                }
                RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "SpreadTheADPResults: after checking number of children");
                TDCOLib.IDCO adpPage = oDoc.GetChild(0);
                String adpPageID = adpPage.ID;
                String adpJsonFileName = BatchPilot.BatchDir + "\\" + adpPageID + "_adp.json";

                RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "SpreadTheADPResults: before reading adp json: " + adpJsonFileName);
                String jsonResults = File.ReadAllText(adpJsonFileName);
                RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "SpreadTheADPResults: after reading adp json");

                if (oDoc.NumOfChildren() == 1)
                {
                    // single-page object, go ahead and apply results to it
                    processADPJson(jsonResults, "", adpPage, adpPageID, 0);
                    utils.getRidOfUnneededADPFields(adpPage, adpConfidenceAction);
                }
                else
                {
                    for (int pageNum = 1; pageNum < oDoc.NumOfChildren(); pageNum++)
                    {
                        TDCOLib.IDCO thisPage = oDoc.GetChild(pageNum);
                        String thisPageID = thisPage.ID;
                        RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "SpreadTheADPResults: processing page " + (pageNum - 1) + " ID " + thisPageID);
                        // apply results to page
                        processADPJson(jsonResults, "", thisPage, thisPageID, pageNum - 1);
                        utils.getRidOfUnneededADPFields(thisPage, adpConfidenceAction);
                    }
                }

                status = true;
            }
            catch (Exception e)
            {
                RRLog.WriteEx(OSADPConnector.LOGGING_ERROR, "error : " + e.Message);
                RRLog.WriteEx(OSADPConnector.LOGGING_ERROR, e.ToString());
                String className = this.GetType().Name;
                String methodName = System.Reflection.MethodInfo.GetCurrentMethod().Name;
                RRLog.WriteEx(OSADPConnector.LOGGING_ERROR, "Class " + className + ", method " + methodName);
                RRLog.WriteEx(OSADPConnector.LOGGING_ERROR, "Stack Trace: " + e.StackTrace);

                status = true; // don't halt rule processing
            }

            RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "***End SpreadTheADPResults");

            return status;
        }

        public bool MoveFieldsToFirstPageOfDocument()
        {
            DateTime compileTime = new DateTime(Builtin.CompileTime, DateTimeKind.Utc); // Ignore the supposed error -- this is resolved at compile time
            RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "***Start MoveFieldsToFirstPageOfDocument - compile time " + compileTime);

            dcSmart.SmartNav smartNav = new dcSmart.SmartNav(this);
            Boolean status = true;

            try
            {
                //Set oPage to the current page object.
                TDCOLib.IDCO oDoc = CurrentDCO;
                string docId = oDoc.ID;

                if (CurrentDCO.ObjectType() != Level.Document)
                {
                    RRLog.WriteEx(OSADPConnector.LOGGING_ERROR, "MoveFieldsToFirstPageOfDocument: Not set at a document level. Returning false.");
                    return false;
                }

                this.utils = new Utils(RRLog);
                // Get the first page in the document
                TDCOLib.IDCO firstPage = GetFirstPageInDoc(oDoc);

                TDCOLib.IDCO thisPage = null;
                int pageNum = 0;
                for (int i = 0; i < oDoc.NumOfChildren(); i++)
                {
                    thisPage = oDoc.GetChild(i);
                    if (thisPage.ObjectType() == Level.Page)
                    {
                        pageNum++;
                        if (!thisPage.ID.Equals(firstPage.ID))
                        {
                            RRLog.WriteEx(OSADPConnector.LOGGING_ERROR, "MoveFieldsToFirstPageOfDocument: Moving fields, etc., for page " + thisPage.ID);
                            addImageFileToThisPageFields(thisPage);
                            MoveFieldsFromPageToPage(thisPage, firstPage, pageNum);
                            SetFirstPageProperty(thisPage, "0");
                            thisPage.Status = 49;
                        }
                        else
                        {
                            SetFirstPageProperty(thisPage, "1");
                            RRLog.WriteEx(OSADPConnector.LOGGING_ERROR, "MoveFieldsToFirstPageOfDocument: Not moving fields, etc., for page " + thisPage.ID + " because it's the fist page of the document");
                        }
                    }
                }

            }
            catch (Exception e)
            {
                RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "error : " + e.Message);
                RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, e.ToString());
                RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, e.StackTrace);
                status = false;
            }

            RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "***End MoveFieldsToFirstPageOfDocument");
            return status;

        }

        private TDCOLib.IDCO GetFirstPageInDoc(TDCOLib.IDCO oDoc)
        {
            TDCOLib.IDCO oPage = null;
            for (int i = 0; i < oDoc.NumOfChildren(); i++)
            {
                oPage = oDoc.GetChild(i);
                if (oPage.ObjectType() == Level.Page)
                {
                    return oPage;
                }
            }
            return oPage;
        }

        private void addImageFileToThisPageFields(TDCOLib.IDCO thisPage, String imageFile = null)
        {
            //String imageFile = null;
            int imageFileVarNum = thisPage.FindVariable("IMAGEFILE");
            if (imageFileVarNum >= 0)
            {
                if (imageFile == null)
                {
                    imageFile = thisPage.GetVariableValue(imageFileVarNum);
                }

                for (int fieldNum = 0; fieldNum < thisPage.NumOfChildren(); fieldNum++)
                {
                    TDCOLib.IDCO thisField = thisPage.GetChild(fieldNum);
                    if (thisField.ObjectType() == Level.Field)
                    {
                        thisField.AddVariable("IMAGEFILE", imageFile);
                        addImageFileToThisPageFields(thisField, imageFile);
                    }
                }
            }

        }

        private void MoveFieldsFromPageToPage(TDCOLib.IDCO fromPage, TDCOLib.IDCO toPage, int pageNum)
        {
            for (int i = 0; i < fromPage.NumOfChildren(); i++)
            {
                TDCOLib.IDCO field = fromPage.GetChild(i);
                if (field.ObjectType() == Level.Field)
                {
                    int imageNameIndex = field.FindVariable("IMAGENAME");
                    String imageName = "";
                    if (imageNameIndex >= 0)
                    {
                        imageName = field.GetVariableValue(imageNameIndex);
                    }
                    MoveFieldFromPlaceToPlace(field, fromPage, toPage, pageNum, imageName);
                }
            }
            // Loop backwards so we can delete as we go
            for (int i = fromPage.NumOfChildren() - 1; i >= 0; i--)
            {
                TDCOLib.IDCO field = fromPage.GetChild(i);
                if (field.ObjectType() == Level.Field)
                {
                    fromPage.DeleteChild(i);
                }
            }
        }

        private void MoveFieldFromPlaceToPlace(TDCOLib.IDCO fieldToMove, TDCOLib.IDCO fromPlace, TDCOLib.IDCO toPlace, int pageNum, String imageName)
        {
            if (fieldToMove.ObjectType() != Level.Field)
            {
                if (fieldToMove.ObjectType() != Level.Character)
                {
                    RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "DocumentCleanUp: MoveFieldFromPlaceToPlace: called for something that is not a field: "
                        + fieldToMove.ID + " type " + fieldToMove.ObjectType() + " text " + fieldToMove.Text);
                }
                return;
            }

            // If we're moving a table
            if (fieldIsATable(fieldToMove))
            {
                // If there's a matching table on the page we're moving to
                TDCOLib.IDCO matchingTable = findMatchingTableField(fieldToMove, toPlace);
                if (matchingTable != null)
                {
                    // Move the line items from this table to the matching table, rather than moving the table to the page
                    moveLineItems(fieldToMove, matchingTable, imageName);
                    return;
                }
            }

            String newFieldID = fieldToMove.ID;
            String newFieldTYPE = fieldToMove.Type;
            String newFieldLabelSuffix = "";
            TDCOLib.IDCO existingField = toPlace.FindChild(newFieldID);
            while (existingField != null)
            {
                StringBuilder newerFieldID = new StringBuilder();
                StringBuilder newerFieldTYPE = new StringBuilder();
                StringBuilder newerFieldLabelSuffix = new StringBuilder();
                newerFieldID.Append(newFieldID).Append("_").Append(pageNum);
                newerFieldTYPE.Append(newFieldTYPE).Append("_").Append(pageNum);
                newerFieldLabelSuffix.Append(newFieldLabelSuffix).Append("_").Append(pageNum);
                RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "DocumentCleanUp: MoveFieldFromPlaceToPlace: found field " + newFieldID + " in new place, trying " + newerFieldID);
                newFieldID = newerFieldID.ToString();
                newFieldTYPE = newerFieldTYPE.ToString();
                newFieldLabelSuffix = newerFieldLabelSuffix.ToString();
                existingField = toPlace.FindChild(newFieldID);
            }

            addFieldFromField(fieldToMove, toPlace, newFieldID, newFieldLabelSuffix, newFieldTYPE, imageName);
        }

        private Boolean fieldIsATable(TDCOLib.IDCO thisField)
        {
            RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "DocumentCleanUp: fieldIsATable: thisField: " + thisField.ID);
            if (thisField.ObjectType() != Level.Field)
            {
                return false;
            }

            Boolean hasSubFields = false;
            Boolean hasSubSubFields = false;
            TDCOLib.IDCO subField = null;
            for (int i = 0; (i < thisField.NumOfChildren()) && !hasSubFields; i++)
            {
                subField = thisField.GetChild(i);
                if (subField.ObjectType() == Level.Field)
                {
                    hasSubFields = true;
                }
            }
            if (hasSubFields)
            {
                for (int i = 0; (i < subField.NumOfChildren()) && !hasSubSubFields; i++)
                {
                    TDCOLib.IDCO subsubField = subField.GetChild(i);
                    if (subsubField.ObjectType() == Level.Field)
                    {
                        hasSubSubFields = true;
                    }
                }
            }
            return hasSubSubFields;
        }

        private TDCOLib.IDCO findMatchingTableField(TDCOLib.IDCO sourceTable, TDCOLib.IDCO pageToLookIn)
        {
            RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "DocumentCleanUp: findMatchingTableField: sourceTable: " + sourceTable.ID + ", pageToLookIn: " + pageToLookIn.ID);
            if (!fieldIsATable(sourceTable))
            {
                return null;
            }
            if (pageToLookIn.ObjectType() != Level.Page)
            {
                return null;
            }
            int entityNameIndex = sourceTable.FindVariable("entityName");
            if (entityNameIndex < 0)
            {
                return null;
            }
            String srcEntityName = sourceTable.GetVariableValue(entityNameIndex);
            if ((srcEntityName == null) || (srcEntityName.Trim().Length == 0))
            {
                return null;
            }
            for (int i = 0; i < pageToLookIn.NumOfChildren(); i++)
            {
                TDCOLib.IDCO thisChild = pageToLookIn.GetChild(i);
                if (fieldIsATable(thisChild))
                {
                    int targetEntityIndex = thisChild.FindVariable("entityName");
                    if (targetEntityIndex >= 0)
                    {
                        String targetEntityName = thisChild.GetVariableValue(targetEntityIndex);
                        if (srcEntityName.Equals(targetEntityName))
                        {
                            return thisChild;
                        }
                    }
                }
            }
            return null;
        }

        private TDCOLib.IDCO getLastLineItem(TDCOLib.IDCO thisTable)
        {
            RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "DocumentCleanUp: getLastLineItem: thisTable: " + thisTable.ID);
            if (!fieldIsATable(thisTable))
            {
                return null;
            }
            for (int i = thisTable.NumOfChildren() - 1; i >= 0; i--)
            {
                TDCOLib.IDCO maybeLastLineItem = thisTable.GetChild(i);
                for (int j = 0; j < maybeLastLineItem.NumOfChildren(); j++)
                {
                    TDCOLib.IDCO lineItemField = maybeLastLineItem.GetChild(j);
                    if (lineItemField.ObjectType() == Level.Field)
                    {
                        return maybeLastLineItem;
                    }
                }
            }
            return null;
        }

        private void moveLineItems(TDCOLib.IDCO fromTable, TDCOLib.IDCO toTable, String imageName)
        {
            RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "DocumentCleanUp: moveLineItems: fromTable: " + fromTable.ID + ", toTable: " + toTable.ID + ", imageName: " + imageName);
            if (!fieldIsATable(fromTable) || !fieldIsATable(toTable))
            {
                throw new Exception("moveLineItems called with non-table items!");
            }
            TDCOLib.IDCO lastToLineItem = getLastLineItem(toTable);
            String pattern = "(\\d+)(?!.*\\d)";
            Match m = Regex.Match(lastToLineItem.ID, pattern);
            int lastLineItemNumber = -1;
            if (m.Success)
            {
                lastLineItemNumber = Int32.Parse(m.Value);
            }
            char[] digits = { '0', '1', '2', '3', '4', '5', '6', '7', '8', '9' };
            String lineItemName = lastToLineItem.ID.Trim(digits);
            for (int i = 0; i < fromTable.NumOfChildren(); i++)
            {
                TDCOLib.IDCO thisLineItem = fromTable.GetChild(i);
                String newFieldID = lineItemName + (lastLineItemNumber + i + 1);
                String newFieldLabelSuffix = "_0";
                addFieldFromField(thisLineItem, toTable, newFieldID, newFieldLabelSuffix, thisLineItem.Type, imageName);
            }
        }

        private void addFieldFromField(TDCOLib.IDCO fieldToMove, TDCOLib.IDCO toPlace, String newFieldID, String newFieldLabelSuffix, String newFieldTYPE, String imageName)
        {
            TDCOLib.IDCO newField = toPlace.AddChild(fieldToMove.ObjectType(), newFieldID, -1);

            for (int i = 0; i < fieldToMove.NumOfVars(); i++)
            {
                String varName = fieldToMove.GetVariableName(i);
                dynamic varValue = fieldToMove.GetVariableValue(i);
                if (varName.ToLower().Trim().Equals("label") || varName.ToLower().Trim().Equals("type"))
                {
                    varValue = varValue + newFieldLabelSuffix;
                }
                if ((varName != null) && (varName.Trim().Length > 0) && (varValue != null))
                {
                    RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "DocumentCleanUp: addFieldFromField: trying to add variable " + varName + " with value " + varValue);
                    try
                    {
                        if (!newField.AddVariable(varName, varValue))
                        {
                            RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "DocumentCleanUp: addFieldFromField: couldn't copy variable " + varName + " with value " + varValue);
                        }
                    }
                    catch (Exception e)
                    {
                        RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "DocumentCleanUp: addFieldFromField: exception " + e.Message + " trying to copy variable " + varName + " with value " + varValue);
                    }
                }
                else
                {
                    RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "DocumentCleanUp: addFieldFromField: couldn't copy variable: name " + varName + ", value " + varValue);

                }
            }
            RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "DocumentCleanUp: addFieldFromField: setting text to " + fieldToMove.Text);
            newField.Text = fieldToMove.Text;
            RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "DocumentCleanUp: addFieldFromField: setting Status to " + fieldToMove.Status);
            newField.Status = fieldToMove.Status;
            RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "DocumentCleanUp: addFieldFromField: setting Type to " + newFieldTYPE);
            newField.Type = newFieldTYPE;
            RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "DocumentCleanUp: addFieldFromField: setting ConfidenceString to " + fieldToMove.ConfidenceString);
            newField.ConfidenceString = fieldToMove.ConfidenceString;
            //newField.ImageName = fieldToMove.ImageName;
            RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "DocumentCleanUp: addFieldFromField: setting Options to " + fieldToMove.Options);
            newField.Options = fieldToMove.Options;
            RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "DocumentCleanUp: addFieldFromField: setting ImageName to " + imageName);
            newField.ImageName = imageName;
            RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "DocumentCleanUp: addFieldFromField: setting IMAGEFILE to " + imageName);
            newField.AddVariable("IMAGEFILE", imageName);

            // Now move the children
            for (int i = 0; i < fieldToMove.NumOfChildren(); i++)
            {
                RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "DocumentCleanUp: addFieldFromField: moving field " + fieldToMove.GetChild(i).ID + " with value " + fieldToMove.GetChild(i).Text);
                MoveFieldFromPlaceToPlace(fieldToMove.GetChild(i), fieldToMove, newField, 0, imageName);
            }

        }

        private void SetFirstPageProperty(TDCOLib.IDCO page, String value)
        {
            String firstPagePropertyName = "IsFirstPage";
            int firstPagePropertyIndex = page.FindVariable(firstPagePropertyName);
            if (firstPagePropertyIndex >= 0)
            {
                page.Variable[firstPagePropertyName] = value;
            }
            else
            {
                page.AddVariable(firstPagePropertyName, value);
            }
        }

    }

}
