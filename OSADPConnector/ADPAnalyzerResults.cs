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
using Newtonsoft.Json.Linq;

namespace OSADPConnector
{
    class ADPAnalyzerResults
    {
        private dynamic ADPJson = null;
        private dclogXLib.IDCLog RRLog = null;
        public ADPAnalyzerResults(String jsonString, dclogXLib.IDCLog RRLog)
        {
            this.RRLog = RRLog;
            try
            {
                ADPJson = JObject.Parse(jsonString);
            }
            catch (Exception e)
            {
                RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "ADPAnalyzerResults: constructor: could not parse ADP JSON results: " + e.Message);
                RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "ADPAnalyzerResults: constructor: stack trace: " + Environment.NewLine + e.StackTrace);
                throw e;
            }
        }

        public String getOCRText()
        {
            dynamic result = null;
            dynamic data = null;
            dynamic DSOutput = null;
            dynamic content = null;
            String ocrText = null;
            try
            {
                result = ADPJson.result[0];
                data = result.data;
                DSOutput = data.DSOutput[0];
                content = DSOutput.Content;
            }
            catch (Exception e)
            {
                RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "ADPAnalyzerResults: getOCRText: could not get OCR Text: " + e.Message);
                RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "ADPAnalyzerResults: getOCRText: stack trace: " + Environment.NewLine + e.StackTrace);
            }
            if (content != null)
            {
                ocrText = (String)content;
            }
            return ocrText;
        }

        public List<Tuple<String, String>> getDocumentClasses()
        {
            dynamic result = null;
            dynamic data = null;
            dynamic Classification = null;
            dynamic DocumentClass = null;
            dynamic Actual = null;
            List<Tuple<String, String>> docClasses = new List<Tuple<String, String>>();
            try
            {
                result = ADPJson.result[0];
                data = result.data;
                Classification = data.Classification;
                DocumentClass = Classification.DocumentClass;
                Actual = DocumentClass.Actual;
                String docClass = (String)Actual;
                String docClassConfidence = "";
                if (DocumentClass.ClassMatch != null) docClassConfidence = DocumentClass.ClassMatch;
                Tuple<String, String> firstClass = new Tuple<String, String>(docClass, docClassConfidence);
                docClasses.Add(firstClass);
                dynamic altDocClassArray = Classification.AlternateDocumentClass;
                for (int i = 0; i < altDocClassArray.Count; i++)
                {
                    dynamic name = altDocClassArray[i].Name;
                    docClass = (String)name;
                    dynamic conf = altDocClassArray[i].ClassMatch;
                    docClassConfidence = (String)conf;
                    Tuple<String, String> thisClass = new Tuple<String, String>(docClass, docClassConfidence);
                    docClasses.Add(thisClass);
                }
            }
            catch (Exception e)
            {
                RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "ADPAnalyzerResults: getDocumentClasses: could not get doc classes: " + e.Message);
                RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "ADPAnalyzerResults: getDocumentClasses: stack trace: " + Environment.NewLine + e.StackTrace);
            }
            return docClasses;
        }

        public List<FieldObject> getBlocks(int pageNum)
        {
            List<FieldObject> resultList = new List<FieldObject>();

            dynamic result = null;
            dynamic data = null;
            dynamic pageList = null;
            dynamic BlockList = null;
            dynamic LineList = null;
            dynamic LineID = null;
            dynamic LineStartX = null;
            dynamic LineStartY = null;
            dynamic LineWidth = null;
            dynamic LineHeight = null;
            dynamic WordList = null;
            dynamic WordID = null;
            dynamic WordStartX = null;
            dynamic WordStartY = null;
            dynamic WordWidth = null;
            dynamic WordHeight = null;
            dynamic WordValue = null;
            dynamic WordOCRConfidence = null;
            dynamic WordCharN = null;
            dynamic bold = null;
            dynamic italic = null;
            dynamic underlined = null;
            dynamic WordFontSize = null;
            dynamic WordFontSizeGroup = null;
            dynamic LineFontFace = null;
            try
            {
                result = ADPJson.result[0];
                data = result.data;
                pageList = data.pageList;
                for (int plIndex = 0; plIndex < pageList.Count; plIndex++)
                {
                    if (plIndex == pageNum)
                    {
                        BlockList = pageList[plIndex].BlockList;
                        for (int blIndex = 0; blIndex < BlockList.Count; blIndex++)
                        {
                            FieldObject blockFO = new FieldObject(FieldObject.FIELD_TYPE_BLOCK, RRLog);
                            dynamic thisBlock = BlockList[blIndex];
                            int left = thisBlock.BlockStartX;
                            int top = thisBlock.BlockStartY;
                            int right = left + (int)thisBlock.BlockWidth;
                            int bottom = top + (int)thisBlock.BlockHeight;
                            PageLocation blockLocation = new PageLocation(left, top, right, bottom);
                            blockFO.setLocation(blockLocation);
                            LineList = thisBlock.LineList;
                            List<String> blockLines = new List<string>();
                            StringBuilder blockQuality = new StringBuilder();
                            for (int llIndex = 0; llIndex < LineList.Count; llIndex++)
                            {
                                StringBuilder lineQuality = new StringBuilder();
                                dynamic thisLine = LineList[llIndex];
                                FieldObject thisLineField = new FieldObject(FieldObject.FIELD_TYPE_LINE, RRLog);
                                left = thisLine.LineStartX;
                                top = thisLine.LineStartY;
                                right = left + (int)thisLine.LineWidth;
                                bottom = top + (int)thisLine.LineHeight;
                                PageLocation lineLocation = new PageLocation(left, top, right, bottom);
                                thisLineField.setLocation(lineLocation);
                                StringBuilder thisLineText = new StringBuilder();
                                WordList = thisLine.WordList;
                                for (int wlIndex = 0; wlIndex < WordList.Count; wlIndex++)
                                {
                                    dynamic thisWord = WordList[wlIndex];
                                    FieldObject thisWordFO = new FieldObject(FieldObject.FIELD_TYPE_WORD, RRLog);
                                    left = thisWord.WordStartX;
                                    top = thisWord.WordStartY;
                                    right = left + (int)thisWord.WordWidth;
                                    bottom = top + (int)thisWord.WordHeight;
                                    PageLocation wordLocation = new PageLocation(left, top, right, bottom);
                                    String wordValue = thisWord.WordValue;
                                    thisLineText.Append(wordValue).Append(" ");
                                    int confidence = getConfidence((String)thisWord.WordOCRConfidence);
                                    lineQuality.Append((String)thisWord.WordOCRConfidence);
                                    int fontSize = 0;
                                    Int32.TryParse((String)thisWord.WordFontSize, out fontSize); ;
                                    thisWordFO.setLocation(wordLocation);
                                    List<String> wordList = new List<string>();
                                    wordList.Add(wordValue);
                                    thisWordFO.setLines(wordList);
                                    thisWordFO.setOriginalLines(wordList);
                                    thisWordFO.setConfidence(confidence);
                                    thisWordFO.setFontSize(fontSize);
                                    thisLineField.add(wordValue, wordLocation, confidence, (String)thisWord.WordFontSize);
                                    thisLineField.add(thisWordFO);
                                }
                                blockQuality.Append(lineQuality.ToString());
                                thisLineField.setConfidence(getConfidence(lineQuality.ToString()));
                                blockLines.Add(thisLineText.ToString());
                                blockFO.add(thisLineField);
                            }
                            blockFO.setConfidence(getConfidence(blockQuality.ToString()));
                            blockFO.setLines(blockLines);
                            blockFO.setOriginalLines(blockLines);
                            resultList.Add(blockFO);
                        }
                    }
                }
            }
            catch (Exception e)
            {
                RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "ADPAnalyzerResults: getLines: could not get lines: " + e.Message);
                RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "ADPAnalyzerResults: getLines: stack trace: " + Environment.NewLine + e.StackTrace);
            }

            return resultList;
        }

        public List<FieldObject> getTables(int pageNum)
        {
            List<FieldObject> resultList = new List<FieldObject>();

            dynamic result = null;
            dynamic data = null;
            dynamic pageList = null;
            dynamic TableList = null;
            dynamic RowList = null;
            dynamic CellList = null;
            dynamic LineList = null;
            dynamic LineID = null;
            dynamic LineStartX = null;
            dynamic LineStartY = null;
            dynamic LineWidth = null;
            dynamic LineHeight = null;
            dynamic WordList = null;
            dynamic WordID = null;
            dynamic WordStartX = null;
            dynamic WordStartY = null;
            dynamic WordWidth = null;
            dynamic WordHeight = null;
            dynamic WordValue = null;
            dynamic WordOCRConfidence = null;
            dynamic WordCharN = null;
            dynamic bold = null;
            dynamic italic = null;
            dynamic underlined = null;
            dynamic WordFontSize = null;
            dynamic WordFontSizeGroup = null;
            dynamic LineFontFace = null;
            try
            {
                result = ADPJson.result[0];
                data = result.data;
                pageList = data.pageList;
                for (int plIndex = 0; plIndex < pageList.Count; plIndex++)
                {
                    if (plIndex == pageNum)
                    {
                        TableList = pageList[plIndex].TableList;
                        for (int blIndex = 0; blIndex < TableList.Count; blIndex++)
                        {
                            FieldObject thisTableFO = new FieldObject(FieldObject.FIELD_TYPE_TABLE, RRLog);
                            dynamic thisTable = TableList[blIndex];
                            int left = thisTable.TableStartX;
                            int top = thisTable.TableStartY;
                            int right = left + (int)thisTable.TableWidth;
                            int bottom = top + (int)thisTable.TableHeight;
                            PageLocation tableLocation = new PageLocation(left, top, right, bottom);
                            thisTableFO.setLocation(tableLocation);
                            RowList = thisTable.RowList;
                            List<String> rowCells = new List<string>();
                            for (int rlIndex = 0; rlIndex < RowList.Count; rlIndex++)
                            {
                                dynamic thisRow = RowList[rlIndex];
                                FieldObject thisRowFO = new FieldObject(FieldObject.FIELD_TYPE_ROW, RRLog);
                                left = thisRow.RowStartX;
                                top = thisRow.RowStartY;
                                right = left + (int)thisRow.RowWidth;
                                bottom = top + (int)thisRow.RowHeight;
                                PageLocation rowLocation = new PageLocation(left, top, right, bottom);
                                thisRowFO.setLocation(rowLocation);
                                CellList = thisRow.CellList;
                                for (int clIndex = 0; clIndex < CellList.Count; clIndex++)
                                {
                                    StringBuilder cellQuality = new StringBuilder();
                                    dynamic thisCell = CellList[clIndex];
                                    FieldObject thisCellFO = new FieldObject(FieldObject.FIELD_TYPE_CELL, RRLog);
                                    left = thisCell.CellStartX;
                                    top = thisCell.CellStartY;
                                    right = left + (int)thisCell.CellWidth;
                                    bottom = top + (int)thisCell.CellHeight;
                                    PageLocation cellLocation = new PageLocation(left, top, right, bottom);
                                    thisCellFO.setLocation(cellLocation);
                                    StringBuilder thisCellText = new StringBuilder();

                                    LineList = thisCell.LineList;
                                    for (int llIndex = 0; llIndex < LineList.Count; llIndex++)
                                    {
                                        StringBuilder lineQuality = new StringBuilder();
                                        dynamic thisLine = LineList[llIndex];
                                        FieldObject thisLineFO = new FieldObject(FieldObject.FIELD_TYPE_LINE, RRLog);
                                        left = thisLine.LineStartX;
                                        top = thisLine.LineStartY;
                                        right = left + (int)thisLine.LineWidth;
                                        bottom = top + (int)thisLine.LineHeight;
                                        PageLocation lineLocation = new PageLocation(left, top, right, bottom);
                                        thisLineFO.setLocation(lineLocation);
                                        StringBuilder thisLineText = new StringBuilder();

                                        WordList = thisLine.WordList;
                                        String cellFontSize = null;
                                        for (int wlIndex = 0; wlIndex < WordList.Count; wlIndex++)
                                        {
                                            dynamic thisWord = WordList[wlIndex];
                                            FieldObject thisWordFO = new FieldObject(FieldObject.FIELD_TYPE_WORD, RRLog);
                                            left = thisWord.WordStartX;
                                            top = thisWord.WordStartY;
                                            right = left + (int)thisWord.WordWidth;
                                            bottom = top + (int)thisWord.WordHeight;
                                            PageLocation wordLocation = new PageLocation(left, top, right, bottom);
                                            String wordValue = thisWord.WordValue;
                                            thisCellText.Append(wordValue).Append(" ");
                                            int confidence = getConfidence((String)thisWord.WordOCRConfidence);
                                            lineQuality.Append((String)thisWord.WordOCRConfidence);
                                            int fontSize = 0;
                                            Int32.TryParse((String)thisWord.WordFontSize, out fontSize); ;
                                            thisWordFO.setLocation(wordLocation);
                                            List<String> wordList = new List<string>();
                                            wordList.Add(wordValue);
                                            thisWordFO.setLines(wordList);
                                            thisWordFO.setOriginalLines(wordList);
                                            thisWordFO.setConfidence(confidence);
                                            thisWordFO.setFontSize(fontSize);
                                            thisLineFO.add(wordValue, wordLocation, confidence, (String)thisWord.WordFontSize);
                                            thisLineFO.add(thisWordFO);
                                            thisLineFO.setFontSize(thisWordFO.getFontSize());
                                            cellFontSize = (String)thisWord.WordFontSize;
                                        }
                                        thisCellFO.add(thisLineFO);
                                        thisCellFO.add(thisLineFO.getAllLines(), thisLineFO.getLocation(), getConfidence(thisCellText.ToString()), cellFontSize);
                                        thisCellFO.setFontSize(thisLineFO.getFontSize());

                                    }
                                    thisRowFO.add(thisCellFO);
                                    thisRowFO.add(thisCellFO.getAllLines(), thisCellFO.getLocation());

                                }
                                thisTableFO.add(thisRowFO);
                                thisTableFO.add(thisRowFO.getAllLines(), thisRowFO.getLocation());

                            }
                            resultList.Add(thisTableFO);
                        }
                    }
                }
            }
            catch (Exception e)
            {
                RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "ADPAnalyzerResults: getTables: could not get lines: " + e.Message);
                RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "ADPAnalyzerResults: getTables: stack trace: " + Environment.NewLine + e.StackTrace);
            }

            return resultList;
        }

        private int getConfidence(String conf)
        {
            // Assumption: conf is a string containing digits 0-9
            int length = conf.Length;
            int total = 0;
            for (int i = 0; i < length; i++)
            {
                int thisConf = 0;
                if (Int32.TryParse(conf.Substring(i, 1), out thisConf))
                {
                    total += thisConf;
                    if (thisConf != 0)
                    {
                        total++; // confidence 9 = 100%, 8 = 90%, etc
                    }
                }
            }
            return (total * 10) / length;
        }

        public ADPPageDimensions getPageDimensions()
        {
            ADPPageDimensions adpPD = new ADPPageDimensions();
            dynamic result = null;
            dynamic data = null;
            dynamic pageList = null;
            dynamic PageInfo = null;
            try
            {
                result = ADPJson.result[0];
                data = result.data;
                pageList = data.pageList;
                dynamic firstPageList = pageList[0];
                PageInfo = firstPageList.PageInfo;
                adpPD.dpiX = PageInfo.dpix;
                adpPD.dpiY = PageInfo.dpiy;
                adpPD.PageHeight = PageInfo.PageHeight;
                adpPD.PageOCRCOnfidence = PageInfo.PageOCRConfidence;
                adpPD.PageWidth = PageInfo.PageWidth;
            }
            catch (Exception e)
            {
                RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "ADPAnalyzerResults: getPageDimensions: could not get OCR Text: " + e.Message);
                RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "ADPAnalyzerResults: getPageDimensions: stack trace: " + Environment.NewLine + e.StackTrace);
            }
            return adpPD;

        }

        public List<ADPRanking> getRankings()
        {
            dynamic result = null;
            dynamic data = null;
            dynamic keyClassRankedListArray = null;
            List<ADPRanking> rankings = new List<ADPRanking>();
            try
            {
                result = ADPJson.result[0];
                data = result.data;
                keyClassRankedListArray = data.KeyClassRankedList;
                for (int i = 0; i < keyClassRankedListArray.Count; i++)
                {
                    ADPRanking ranking = new ADPRanking();
                    dynamic kcrlItem = keyClassRankedListArray[i];
                    dynamic kcID = kcrlItem.KeyClassID;
                    ranking.keyClassID = (String)kcID;
                    dynamic kcName = kcrlItem.KeyClassName;
                    ranking.keyClassName = (String)kcName;
                    dynamic kcType = kcrlItem.KeyClassType;
                    ranking.keyClassType = (String)kcType;
                    ranking.kvpRankedList = new List<ADPRank>();
                    dynamic rankedListArray = kcrlItem.KVPRankedList;
                    for (int j = 0; j < rankedListArray.Count; j++)
                    {
                        ADPRank rank = new ADPRank();
                        dynamic pageNo = rankedListArray[j].PageNo;
                        rank.PageNo = (int)pageNo;
                        dynamic kvpid = rankedListArray[j].KVPID;
                        rank.kvpID = (String)kvpid;
                        dynamic reserved = rankedListArray[j].Reserved1;
                        rank.reserved = (String)reserved;
                        ranking.kvpRankedList.Add(rank);
                    }
                    rankings.Add(ranking);
                }
                for (int i = 0; i < rankings.Count; i++)
                {
                    ADPRanking ranking = rankings[i];
                    RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "ADPAnalyzerResults: getRankings: ranking " + i + ": key class name: " + ranking.keyClassName);
                    RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "                                 key class ID: " + ranking.keyClassID);
                    RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "                                 key class type: " + ranking.keyClassType);
                    for (int j = 0; j < ranking.kvpRankedList.Count; j++)
                    {
                        ADPRank rank = ranking.kvpRankedList[j];
                        RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "                                 KVP " + j + " ID: " + rank.kvpID);
                        RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "                                 KVP " + j + " value: " + rank.reserved);
                    }
                }
            }
            catch (Exception e)
            {
                RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "ADPAnalyzerResults: getRankings: exception: could not get key class ranked list: " + e.Message);
                RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "ADPAnalyzerResults: getRankings: stack trace: " + Environment.NewLine + e.StackTrace);
            }
            return rankings;
        }

        public List<ADPKeyValuePair> getKeyValuePairs(List<ADPRanking> adpRankings, int adpFieldsAction, int pageNum)
        {
            RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "***Begin ADPAnalyzerResults: getKeyValuePairs, pageNum: " + pageNum);
            List<ADPKeyValuePair> kvpList = new List<ADPKeyValuePair>();
            dynamic result = null;
            dynamic data = null;
            dynamic pageList = null;
            dynamic KVPTable = null;
            try
            {
                result = ADPJson.result[0];
                data = result.data;
                pageList = data.pageList;
                int lastPage = pageList.Count;
                for (int i = 0; i < lastPage; i++)
                {
                    if (i == pageNum)
                    {
                        dynamic firstPageList = pageList[i];
                        KVPTable = firstPageList.KVPTable;
                        for (int kvpIndex = 0; kvpIndex < KVPTable.Count; kvpIndex++)
                        {
                            dynamic thisKVP = KVPTable[kvpIndex];
                            ADPKeyValuePair kvp = new ADPKeyValuePair();
                            kvp.added = false;
                            if (thisKVP.KeyClass != null) kvp.keyClass = thisKVP.KeyClass;
                            if (thisKVP.KeyClassConfidence != null)
                            {
                                try
                                {
                                    kvp.keyClassConfidence = thisKVP.KeyClassConfidence;
                                    if (kvp.keyClassConfidence.Length == 0)
                                    {
                                        kvp.keyClassConfidence = "Low";
                                    }
                                }
                                catch (Exception e)
                                {
                                    try
                                    {
                                        int confInt = thisKVP.KeyClassConfidence;
                                        if (confInt >= 80)
                                        {
                                            kvp.keyClassConfidence = "High";
                                        }
                                        else if (confInt >= 60)
                                        {
                                            kvp.keyClassConfidence = "Medium";
                                        }
                                        else
                                        {
                                            kvp.keyClassConfidence = "Low";
                                        }
                                    }
                                    catch (Exception eInner)
                                    {
                                        try
                                        {
                                            float confFloat = thisKVP.KeyClassConfidence;
                                            if (confFloat >= 80.0)
                                            {
                                                kvp.keyClassConfidence = "High";
                                            }
                                            else if (confFloat >= 60.0)
                                            {
                                                kvp.keyClassConfidence = "Medium";
                                            }
                                            else
                                            {
                                                kvp.keyClassConfidence = "Low";
                                            }
                                        }
                                        catch (Exception eNoneOfAbove)
                                        {
                                            RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "ADPAnalyzerResults: getKeyValuePairs: kvp has invalid KeyClassConfidence! " + thisKVP);
                                        }
                                    }
                                }
                            }
                            int confIntString = 0;
                            if (Int32.TryParse(kvp.keyClassConfidence, out confIntString))
                            {
                                if (confIntString >= 80)
                                {
                                    kvp.keyClassConfidence = "High";
                                }
                                else if (confIntString >= 60)
                                {
                                    kvp.keyClassConfidence = "Medium";
                                }
                                else
                                {
                                    kvp.keyClassConfidence = "Low";
                                }
                            }
                            try
                            {
                                float confFloatString = float.Parse(kvp.keyClassConfidence);
                                if (confFloatString >= 80.0)
                                {
                                    kvp.keyClassConfidence = "High";
                                }
                                else if (confFloatString >= 60.0)
                                {
                                    kvp.keyClassConfidence = "Medium";
                                }
                                else
                                {
                                    kvp.keyClassConfidence = "Low";
                                }
                            }
                            catch (Exception e) { }
                            if (thisKVP.Key != null) kvp.key = thisKVP.Key;
                            if (thisKVP.Value != null) kvp.value = thisKVP.Value;
                            if (thisKVP.ValueStartX != null) kvp.valueX = thisKVP.ValueStartX;
                            if (thisKVP.ValueStartY != null) kvp.valueY = thisKVP.ValueStartY;
                            if (thisKVP.ValueWidth != null) kvp.valueWidth = thisKVP.ValueWidth;
                            if (thisKVP.ValueHeight != null) kvp.valueHeight = thisKVP.ValueHeight;
                            if (thisKVP.Sensitivity != null) kvp.sensitivity = thisKVP.Sensitivity;
                            if (thisKVP.ValueConfidence != null) kvp.confidence = ((int)thisKVP.ValueConfidence + 1) * 10;
                            if ((thisKVP.Value != null) && !((String)thisKVP.ValueType).ToLower().Equals("table"))
                            {
                                try
                                {
                                    if (thisKVP.KeyStartX != null) kvp.keyX = thisKVP.KeyStartX;
                                    if (thisKVP.KeyStartY != null) kvp.keyY = thisKVP.KeyStartY;
                                    if (thisKVP.KeyWidth != null) kvp.keyWidth = thisKVP.KeyWidth;
                                    if (thisKVP.KeyHeight != null) kvp.keyHeight = thisKVP.KeyHeight;
                                    if (thisKVP.OriginalKey != null) kvp.originalKey = thisKVP.OriginalKey;
                                    if (thisKVP.OriginalValue != null) kvp.originalValue = thisKVP.OriginalValue;
                                }
                                catch (Exception e)
                                {
                                    RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "ADPAnalyzerResults: getKeyValuePairs: kvp has invalid KeyStartX, KeyStartY, KeyWidth, KeyHeight, OriginalKey, or OriginalValue! " + thisKVP);
                                }
                            }
                            else
                            {
                                // Have table in KVPTable
                                // 
                            }
                            if (thisKVP.KVPID != null) kvp.kvpID = thisKVP.KVPID;
                            if (thisKVP.KeyClass != null) kvp.keyClass = thisKVP.KeyClass;
                            if (thisKVP.KeyClassID != null) kvp.keyClassID = thisKVP.KeyClassID;

                            RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "ADPAnalyzerResults: getKeyValuePairs: examining KVP, key class: " + kvp.keyClass
                                + ", key: " + kvp.key + ", value: " + kvp.value + ", ID: " + kvp.kvpID);
                            Boolean addTheKVP = true;
                            //xxx handle no doctype and ADP fields action from parms
                            int fieldActionToUse = adpFieldsAction;
                            // Should we keep all KVPs?
                            if (fieldActionToUse == OSADPConnector.ADPFIELDS_KEEPALL)
                            {
                                addTheKVP = true;
                                RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "ADPAnalyzerResults: getKeyValuePairs: ka: no key class so not including this key");
                            }
                            // Should we require KVPs to have a key class?
                            else if (fieldActionToUse == OSADPConnector.ADPFIELDS_KEEPALLWITHKEYCLASS)
                            {
                                if ((kvp.keyClass == null) || (kvp.keyClass.Length == 0))
                                {
                                    addTheKVP = false;
                                    RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "ADPAnalyzerResults: getKeyValuePairs: kawkc: no key class so not including this key");
                                }
                            }
                            // Should we be keeping just the single best instance (as long as it has a key class)?
                            else if (fieldActionToUse == OSADPConnector.ADPFIELDS_KEEPSINGLEBESTWITHKEYCLASS)
                            {
                                if ((kvp.keyClass == null) || (kvp.keyClass.Length == 0))
                                {
                                    addTheKVP = false;
                                    RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "ADPAnalyzerResults: getKeyValuePairs: ksbwkc: no key class so not including this key");
                                }
                                else // Find this kvp in the kvp ranking
                                {
                                    int indexInRanking = findRanking(adpRankings, kvp.keyClass, kvp.keyClassID, kvp.kvpID);
                                    if (indexInRanking == 0)
                                    {
                                        addTheKVP = true;
                                        RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "ADPAnalyzerResults: getKeyValuePairs: ksbwkc: kvp was first in ranking, keeping it");
                                    }
                                    else if (indexInRanking < 0)
                                    {
                                        addTheKVP = false;
                                        RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "ADPAnalyzerResults: getKeyValuePairs: ksbwkc: couldn't find kvp in ranking area, not including it");
                                    }
                                    else
                                    {
                                        addTheKVP = false;
                                        RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "ADPAnalyzerResults: getKeyValuePairs: ksbwkc: kvp was not first in ranking (was number " + indexInRanking +
                                            ", counting from 0, so not including it");
                                    }
                                }
                            }
                            // Should we keep the single best instance (whether or not it has a key class)?
                            else if (fieldActionToUse == OSADPConnector.ADPFIELDS_KEEPSINGLEBEST)
                            {
                                if ((kvp.keyClass == null) || (kvp.keyClass.Length == 0))
                                {
                                    addTheKVP = true;
                                    RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "ADPAnalyzerResults: getKeyValuePairs: ksb: kvp was null or empty, keeping it");
                                }
                                else // Find this kvp in the kvp ranking
                                {
                                    int indexInRanking = findRanking(adpRankings, kvp.keyClass, kvp.keyClassID, kvp.kvpID);
                                    if (indexInRanking == 0)
                                    {
                                        addTheKVP = true;
                                        RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "ADPAnalyzerResults: getKeyValuePairs: ksb: kvp was 1st in ranking, keeping it");
                                    }
                                    else if (indexInRanking < 0)
                                    {
                                        addTheKVP = true;
                                        RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "ADPAnalyzerResults: getKeyValuePairs: ksb: couldn't find kvp in ranking area, including it");
                                    }
                                    else
                                    {
                                        addTheKVP = false;
                                        RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "ADPAnalyzerResults: getKeyValuePairs: ksb: kvp wasn't first in ranking (was number " + indexInRanking +
                                            ", counting from 0, so not including it");
                                    }
                                }
                            }
                            if (addTheKVP)
                            {
                                kvpList.Add(kvp);
                            }
                        }
                    }
                }
            }
            catch (Exception e)
            {
                RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "ADPAnalyzerResults: getKeyValuePairs: could not get OCR Text: " + e.Message);
                RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "ADPAnalyzerResults: getKeyValuePairs: stack trace: " + Environment.NewLine + e.StackTrace);
            }
            RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "***End ADPAnalyzerResults: getKeyValuePairs");
            RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "ADPAnalyzerResults: getKeyValuePairs: LogKVPList begin");
            LogKVPList(kvpList, "");
            RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "ADPAnalyzerResults: getKeyValuePairs: LogKVPList end");
            return kvpList;

        }

        private int findRanking(List<ADPRanking> adpRankings, String keyClass, String keyClassID, String kvpID)
        {
            foreach (ADPRanking adpRanking in adpRankings)
            {
                if ((adpRanking.keyClassID != null) && adpRanking.keyClassID.Equals(keyClassID) && (adpRanking.keyClassName != null) && adpRanking.keyClassName.Equals(keyClass))
                {
                    for (int i = 0; i < adpRanking.kvpRankedList.Count; i++)
                    {
                        ADPRank rank = adpRanking.kvpRankedList[i];
                        if ((rank.kvpID != null) && rank.kvpID.Equals(kvpID))
                        {
                            return i;
                        }
                    }
                    return -1;
                }
            }
            return -1;
        }


        public List<ADPKeyValuePair> getTableKeyValuePairs(int pageNum)
        {
            RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "***Begin ADPAnalyzerResults: getTableKeyValuePairs");
            List<ADPKeyValuePair> kvpList = new List<ADPKeyValuePair>();
            dynamic result = null;
            dynamic data = null;
            dynamic pageList = null;
            dynamic TabularBlockDiscovery = null;
            dynamic KVPTable = null;
            try
            {
                result = ADPJson.result[0];
                data = result.data;
                pageList = data.pageList;
                int lastPage = pageList.Count;
                for (int i = 0; i < lastPage; i++)
                {
                    if (i == pageNum)
                    {
                        dynamic firstPageList = pageList[i];
                        KVPTable = firstPageList.KVPTable;
                        //}
                        for (int kvpIndex = 0; kvpIndex < KVPTable.Count; kvpIndex++)
                        {
                            dynamic thisKVP = KVPTable[kvpIndex];
                            dynamic tableJson = null;
                            if (!((String)thisKVP.ValueType).ToLower().Equals("table"))
                            {
                                RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "ADPAnalyzerResults: getTableKeyValuePairs: key " + thisKVP.Key + " is not ValueType table -- skipping it");
                                continue;
                            }
                            try
                            {
                                tableJson = thisKVP.ComplexKVPStructure;
                            }
                            catch (Exception e)
                            {
                                // do nothing
                            }
                            if (tableJson == null)
                            {
                                RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "ADPAnalyzerResults: getTableKeyValuePairs: key " + thisKVP.Key + " has no ComplexKVPStructure -- skipping it");
                                continue;
                            }
                            ADPKeyValuePair adpKVPTable = new ADPKeyValuePair();
                            if (thisKVP.KeyClass != null) adpKVPTable.keyClass = thisKVP.KeyClass;
                            if (thisKVP.KeyClassConfidence != null)
                            {
                                if (thisKVP.KeyClassConfidence is String)
                                {
                                    adpKVPTable.keyClassConfidence = thisKVP.KeyClassConfidence;
                                    if (adpKVPTable.keyClassConfidence.Length == 0)
                                    {
                                        adpKVPTable.keyClassConfidence = "Low";
                                    }
                                }
                                else if (thisKVP.KeyClassConfidence is Int32)
                                {
                                    if (thisKVP.KeyClassConfidence >= 80)
                                    {
                                        adpKVPTable.keyClassConfidence = "High";
                                    }
                                    else if (thisKVP.KeyClassConfidence >= 60)
                                    {
                                        adpKVPTable.keyClassConfidence = "Medium";
                                    }
                                    else
                                    {
                                        adpKVPTable.keyClassConfidence = "Low";
                                    }
                                }
                                else if (thisKVP.KeyClassConfidence is float)
                                {
                                    if (thisKVP.KeyClassConfidence >= 80.0)
                                    {
                                        adpKVPTable.keyClassConfidence = "High";
                                    }
                                    else if (thisKVP.KeyClassConfidence >= 60.0)
                                    {
                                        adpKVPTable.keyClassConfidence = "Medium";
                                    }
                                    else
                                    {
                                        adpKVPTable.keyClassConfidence = "Low";
                                    }
                                }
                            }
                            if (thisKVP.Key != null) adpKVPTable.key = thisKVP.Key;
                            if (thisKVP.Value != null) adpKVPTable.value = thisKVP.Value;
                            if (thisKVP.ValueStartX != null) adpKVPTable.valueX = getIntValue(thisKVP.ValueStartX);
                            if (thisKVP.ValueStartY != null) adpKVPTable.valueY = getIntValue(thisKVP.ValueStartY);
                            if (thisKVP.ValueWidth != null) adpKVPTable.valueWidth = getIntValue(thisKVP.ValueWidth);
                            if (thisKVP.ValueHeight != null) adpKVPTable.valueHeight = getIntValue(thisKVP.ValueHeight);
                            if (thisKVP.Sensitivity != null) adpKVPTable.sensitivity = thisKVP.Sensitivity;
                            if ((thisKVP.Value != null) && !((String)thisKVP.Value).Equals("_TABLE_ZONE_"))
                            {
                                if (thisKVP.KeyStartX != null) adpKVPTable.keyX = getIntValue(thisKVP.KeyStartX);
                                if (thisKVP.KeyStartY != null) adpKVPTable.keyY = getIntValue(thisKVP.KeyStartY);
                                if (thisKVP.KeyWidth != null) adpKVPTable.keyWidth = getIntValue(thisKVP.KeyWidth);
                                if (thisKVP.KeyHeight != null) adpKVPTable.keyHeight = getIntValue(thisKVP.KeyHeight);
                                if (thisKVP.OriginalKey != null) adpKVPTable.originalKey = thisKVP.OriginalKey;
                                if (thisKVP.OriginalValue != null) adpKVPTable.originalValue = thisKVP.OriginalValue;
                            }

                            try
                            {
                                // Have main table, now need line items
                                dynamic ComplexKVPStructure = thisKVP.ComplexKVPStructure;
                                RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "ADPAnalyzerResults: getTableKeyValuePairs: key " + thisKVP.Key + " has ComplexKVPStructure");
                                dynamic Attributes = ComplexKVPStructure.Attributes;
                                for (int aIndex = 0; aIndex < Attributes.Count; aIndex++)
                                {
                                    dynamic ValueList = Attributes[aIndex].ValueList;
                                    for (int vlIndex = 0; vlIndex < ValueList.Count; vlIndex++)
                                    {
                                        // Now have line item
                                        ADPKeyValuePair lineItem = new ADPKeyValuePair();

                                        // Populate line item with attribute KVPs
                                        dynamic InnerComplexKVPStructure = ValueList[vlIndex].ComplexKVPStructure;
                                        dynamic InnerAttributes = InnerComplexKVPStructure.Attributes;
                                        dynamic sourceLineItem = ValueList[vlIndex];
                                        if (sourceLineItem.ValueStartX != null) lineItem.valueX = getIntValue(sourceLineItem.ValueStartX);
                                        if (sourceLineItem.ValueStartY != null) lineItem.valueY = getIntValue(sourceLineItem.ValueStartY);
                                        if (sourceLineItem.ValueWidth != null) lineItem.valueWidth = getIntValue(sourceLineItem.ValueWidth);
                                        if (sourceLineItem.ValueHeight != null) lineItem.valueHeight = getIntValue(sourceLineItem.ValueHeight);
                                        if (sourceLineItem.LineItemID != null) lineItem.LineItemID = getIntValue(sourceLineItem.LineItemID);
                                        if (sourceLineItem.SeqLineItemID != null) lineItem.SeqLineItemID = getIntValue(sourceLineItem.SeqLineItemID);
                                        if (sourceLineItem.Sensitivity != null) lineItem.sensitivity = sourceLineItem.Sensitivity;
                                        lineItem.haveLineItem = false;
                                        if ((sourceLineItem.LineItemID != null) && (sourceLineItem.SeqLineItemID != null)) lineItem.haveLineItem = true;
                                        for (int iaIndex = 0; iaIndex < InnerAttributes.Count; iaIndex++)
                                        {
                                            dynamic attribute = InnerAttributes[iaIndex];
                                            RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "ADPAnalyzerResults: getTableKeyValuePairs: lineitem " + attribute.Key + ": " + attribute.Value);
                                            String checkLineItemKey = attribute.Key;
                                            String checkLineItemValue = attribute.Value;
                                            if ((checkLineItemKey != null) && (checkLineItemKey.Length > 0) && (checkLineItemValue != null) && (checkLineItemValue.Length > 0))
                                            {
                                                // Have cell source
                                                ADPKeyValuePair cell = new ADPKeyValuePair();

                                                // Populate cell info
                                                if (attribute.KeyClass != null) cell.keyClass = attribute.KeyClass;
                                                if (attribute.KeyClassConfidence != null) cell.keyClassConfidence = attribute.KeyClassConfidence;
                                                if (attribute.Key != null) cell.key = attribute.Key;
                                                if (attribute.Value != null) cell.value = attribute.Value;
                                                if (attribute.ValueStartX != null) cell.valueX = getIntValue(attribute.ValueStartX, "table row " + aIndex + ", attribute " + iaIndex + ", ValueStartX");
                                                if (attribute.ValueStartY != null) cell.valueY = getIntValue(attribute.ValueStartY, "table row " + aIndex + ", attribute " + iaIndex + ", ValueStartY");
                                                if (attribute.ValueWidth != null) cell.valueWidth = getIntValue(attribute.ValueWidth, "table row " + aIndex + ", attribute " + iaIndex + ", ValueWidth");
                                                if (attribute.ValueHeight != null) cell.valueHeight = getIntValue(attribute.ValueHeight, "table row " + aIndex + ", attribute " + iaIndex + ", ValueHeight");
                                                if (attribute.KeyStartX != null) cell.keyX = getIntValue(attribute.KeyStartX, "table row " + aIndex + ", attribute " + iaIndex + ", KeyStartX");
                                                if (attribute.KeyStartY != null) cell.keyY = getIntValue(attribute.KeyStartY, "table row " + aIndex + ", attribute " + iaIndex + ", KeyStartY");
                                                if (attribute.KeyWidth != null) cell.keyWidth = getIntValue(attribute.KeyWidth, "table row " + aIndex + ", attribute " + iaIndex + ", KeyWidth");
                                                if (attribute.KeyHeight != null) cell.keyHeight = getIntValue(attribute.KeyHeight, "table row " + aIndex + ", attribute " + iaIndex + ", KeyHeight");
                                                if (attribute.OriginalKey != null) cell.originalKey = attribute.OriginalKey;
                                                if (attribute.OriginalValue != null) cell.originalValue = attribute.OriginalValue;
                                                if (attribute.ValueConfidence != null) cell.confidence = ((int)getIntValue(attribute.ValueConfidence) + 1) * 10;
                                                if (attribute.Sensitivity != null) cell.sensitivity = attribute.Sensitivity;

                                                // Add cell to line item
                                                lineItem.nested.Add(cell);
                                                RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "ADPAnalyzerResults: getTableKeyValuePairs: just added nested item " +
                                                    cell.key + "(" + cell.value + ") to key " + thisKVP.Key);
                                            }
                                            else
                                            {
                                                RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "ADPAnalyzerResults: getTableKeyValuePairs: skipping empty nested item");
                                            }
                                        }

                                        // Add line item to table
                                        adpKVPTable.nested.Add(lineItem);
                                    }
                                }
                            }
                            catch (Exception e)
                            {
                                RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "ADPAnalyzerResults: getTableKeyValuePairs: exception: probably not a problem, but no nested table for key " + thisKVP.Key + ": " + e.Message);
                                RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "ADPAnalyzerResults: getTableKeyValuePairs: stack trace: " + Environment.NewLine + e.StackTrace);
                            }
                            kvpList.Add(adpKVPTable);
                        }
                    }
                }
            }
            catch (Exception e)
            {
                RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "ADPAnalyzerResults: getTableKeyValuePairs: could not get OCR Text: " + e.Message);
                RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "ADPAnalyzerResults: getTableKeyValuePairs: stack trace: " + Environment.NewLine + e.StackTrace);
            }
            RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "***End ADPAnalyzerResults: getTableKeyValuePairs");
            RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "ADPAnalyzerResults: getTableKeyValuePairs: LogKVPList begin");
            LogKVPList(kvpList, "");
            RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "ADPAnalyzerResults: getTableKeyValuePairs: LogKVPList end");
            return kvpList;

        }

        private int getIntValue(dynamic input, String message = null)
        {
            int intValue = -1;
            try
            {
                intValue = input;
            }
            catch (Exception e)
            {
                if (message != null)
                {
                    RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "ADPAnalyzerResults: getIntValue: unassignable value for " + message + ": " + input);
                }
                RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "ADPAnalyzerResults: getIntValue: exception: " + e.Message);

            }
            return intValue;
        }

        public void LogKVPList(List<ADPKeyValuePair> kvpList, String message)
        {
            if ((kvpList != null) && (kvpList.Count > 0))
            {
                foreach (ADPKeyValuePair kvp in kvpList)
                {
                    // Log the header info
                    RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, message + kvp.keyClass + ": " + kvp.value +
                        " key loc (" + kvp.keyX + ", " + kvp.keyY + ", " + (kvp.keyX + kvp.keyWidth) + ", " + (kvp.keyY + kvp.keyHeight) + ") " +
                        " value loc (" + kvp.valueX + ", " + kvp.valueY + ", " + (kvp.valueX + kvp.valueWidth) + ", " + (kvp.valueY + kvp.valueHeight) + ") " +
                        " key width/height (" + kvp.keyWidth + ", " + kvp.keyHeight + ") value width/height (" + kvp.valueWidth + ", " + kvp.valueHeight + ")");

                    // Log the nested info
                    LogKVPList(kvp.nested, message + "  ");
                }
            }
        }

    }
}
