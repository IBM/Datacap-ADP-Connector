//
// © Copyright IBM Corp. 1994, 2023 All Rights Reserved
//
// Created by Scott Sumner-Moore, 2023
//


using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.Serialization.Json;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;

namespace OSADPConnector
{
    class Utils
    {
        private dclogXLib.IDCLog RRLog;

        public Utils(dclogXLib.IDCLog RRLog)
        {
            this.RRLog = RRLog;
        }

        public static ADPConfig loadADPConfig(String config, dclogXLib.IDCLog RRLog)
        {
            DataContractJsonSerializer js = new DataContractJsonSerializer(typeof(ADPConfig));
            byte[] configBuffer = Encoding.ASCII.GetBytes(config);
            MemoryStream msObj = new MemoryStream(configBuffer);
            ADPConfig adpConfig = (ADPConfig)js.ReadObject(msObj);
            fixAnalyzeTarget(adpConfig);
            RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "loadADPConfigString: config: " + config);
            RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "loadADPConfigString: adpConfig: " + adpConfig);

            return adpConfig;
        }

        public void getRidOfUnneededADPFields(TDCOLib.IDCO oPage, int adpConfidenceAction = OSADPConnector.ADP_KEEPALL)
        {
            RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "***Start getRidOfUnneededADPFields");
            int adpIndicatorToUse = adpConfidenceAction;
            // Now get rid of ADP entries we shouldn't have
            String actionString = null;
            if (adpIndicatorToUse == OSADPConnector.ADP_DELETEALL)
            {
                actionString = "delete all";
            }
            else if (adpIndicatorToUse == OSADPConnector.ADP_KEEPALL)
            {
                actionString = "keep all";
            }
            else if (adpIndicatorToUse == OSADPConnector.ADP_KEEPHIGH)
            {
                actionString = "keep only high confidence";
            }
            else if (adpIndicatorToUse == OSADPConnector.ADP_KEEPMEDIUM)
            {
                actionString = "keep high and medium confidence";
            }

            RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "AnnotatorAction: getRidOfUnneededADPFields: we will " + actionString + " ADP fields");
            if (adpIndicatorToUse != OSADPConnector.ADP_KEEPALL)
            {
                for (int childIndex = oPage.NumOfChildren(); childIndex >= 0; childIndex--)
                {
                    TDCOLib.IDCO pageField = oPage.GetChild(childIndex);
                    int entityTypeIndex = -1;
                    if (pageField != null)
                    {
                        entityTypeIndex = pageField.FindVariable("entityType");
                    }
                    if (entityTypeIndex >= 0)
                    {
                        String entityTypeValue = pageField.GetVariableValue(entityTypeIndex);
                        if ((entityTypeValue != null) && entityTypeValue.Equals("ADP"))
                        {
                            Boolean deleteIt = false;
                            String adpKeyClassConfidence = null;
                            if (adpIndicatorToUse == OSADPConnector.ADP_DELETEALL)
                            {
                                deleteIt = true;
                            }
                            else
                            {
                                int adpKeyClassConfIndex = pageField.FindVariable(ADPKeyValuePair.ADPKeyClassConfidence);
                                if (adpKeyClassConfIndex >= 0)
                                {
                                    adpKeyClassConfidence = pageField.GetVariableValue(adpKeyClassConfIndex);

                                    if ((adpIndicatorToUse == OSADPConnector.ADP_KEEPHIGH) && !adpKeyClassConfidence.ToLower().Trim().Equals("high"))
                                    {
                                        deleteIt = true;
                                    }
                                    else if ((adpIndicatorToUse == OSADPConnector.ADP_KEEPMEDIUM) && adpKeyClassConfidence.ToLower().Trim().Equals("low"))
                                    {
                                        deleteIt = true;
                                    }
                                }
                            }
                            if (deleteIt)
                            {
                                int keyPosIndex = pageField.FindVariable("KeyPosition");
                                String keyPosition = null;
                                if (keyPosIndex >= 0)
                                {
                                    keyPosition = pageField.GetVariableValue(keyPosIndex);
                                }
                                int valuePosIndex = pageField.FindVariable("Position");
                                String valuePosition = null;
                                if (valuePosIndex >= 0)
                                {
                                    valuePosition = pageField.GetVariableValue(valuePosIndex);
                                }
                                int idIndex = pageField.FindVariable("ID");
                                String idFromField = null;
                                if (idIndex >= 0)
                                {
                                    idFromField = pageField.GetVariableValue(idIndex);
                                }
                                int keyIndex = pageField.FindVariable("KeyMatch");
                                String key = null;
                                if (keyIndex >= 0)
                                {
                                    key = pageField.GetVariableValue(keyIndex);
                                }
                                int valueIndex = pageField.FindVariable("subMatch1");
                                String value = null;
                                if (valueIndex >= 0)
                                {
                                    value = pageField.GetVariableValue(valueIndex);
                                }
                                RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "AnnotatorAction: getRidOfUnneededADPFields: deleting ADP field " + childIndex + ", ID: " + idFromField
                                    + ", key: (" + keyPosition + ") " + key + ", value: (" + valuePosition + ") " + value);
                                oPage.DeleteChild(childIndex);
                            }
                        }
                    }
                }
            }
            RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "***End getRidOfUnneededADPFields");
        }

        private static void fixAnalyzeTarget(ADPConfig adpConfig)
        {
            if (adpConfig.analyze_target.Contains("[[adp_project_id]]"))
            {
                int leftEdge = adpConfig.analyze_target.IndexOf("[[");
                int rightEdge = adpConfig.analyze_target.IndexOf("]]") + 2;
                StringBuilder atStringBuilder = new StringBuilder();
                atStringBuilder.Append(adpConfig.analyze_target.Substring(0, leftEdge));
                atStringBuilder.Append(adpConfig.adp_project_id);
                atStringBuilder.Append(adpConfig.analyze_target.Substring(rightEdge));
                adpConfig.analyze_target = atStringBuilder.ToString();
            }
        }

        public void updateKeyMatchEntityAndProperties(TDCOLib.IDCO oPage, /*IEntity entity,*/ String name, String labelName, String typeName, String text, PageLocation location, int validityPercentage, String altName,
            Double similarity = -1.0, String altType = null)
        {

            if (altName != null) // In some cases we need a different name, specifically if display_all is set, this will allow field creation of the name
            {
                labelName = altName;
            }

            RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "***Start Utils: updateKeyMatchEntityAndProperties");
            RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "updateKeyMatchEntityAndProperties: label: " + labelName);

            //String typeName = entity.getLabel();
            if (altType != null)
            {
                typeName = altType;
            }

            TDCOLib.IDCO oField;

            oField = oPage.FindChild(labelName);
            if (oField == null) // <--- This is an issue for tables
            {
                RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "updateKeyMatchEntityAndProperties: Entity: " + labelName + " not found on Datacap Page, creating");
                oField = oPage.AddChild(3, labelName, -1);
            }


            if (oField != null)
            {
                oField.Status = 0;
                RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "updateKeyMatchEntityAndProperties: status: 0");
                oField.Type = typeName;
                oField.set_Variable("entityType", typeName);
                oField.set_Variable("Position", location.getPosition());
                RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "updateKeyMatchEntityAndProperties: Type: " + oField.Type + " Position: " + location.getPosition());
                oField.set_Variable("l", location.getLeft() + "");
                oField.set_Variable("t", location.getTop() + "");
                oField.set_Variable("r", location.getRight() + "");
                oField.set_Variable("b", location.getBottom() + "");
                oField.set_Variable("entityName", name);

                oField.set_Variable("KeyMatch", text);
                oField.set_Variable("KeyPosition", location.getPosition());

                if (similarity > 0.0)
                {
                    RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "updateKeyMatchEntityAndProperties: key similarity: " + similarity.ToString());
                    oField.set_Variable("KeySimilarity", similarity.ToString());
                }
            }


            RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "***End Utils: updateKeyMatchEntityAndProperties");
        }

        public void updateEntityAndProperties(TDCOLib.IDCO oPage, /*IEntity entity,*/ String name, String labelName, String typeName, String text, PageLocation location, int validityPercentage, String altName,
            IDictionary<String, String> extraData, FieldObject page, Double similarity = -1.0, String altType = null)
        {
            if (altName != null) // In some cases we need a different name, specifically if display_all is set, this will allow field creation of the name
            {
                labelName = altName;
            }

            RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "***Start Utils: updateEntityAndProperties");
            RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "updateEntityAndProperties: label: " + labelName);


            TDCOLib.IDCO oField;

            oField = oPage.FindChild(labelName);
            if (oField == null)
            {
                RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "updateEntityAndProperties: Entity: " + labelName + " not found on Datacap Page, creating");
                oField = oPage.AddChild(3, labelName, -1);
            }


            if (oField != null)
            {
                if (page != null) // Testing for the page, if the clavis is passed in via a ruleset the page in not found, so trust the validityPercentage passed in initially
                {
                    validityPercentage = decreaseValidityPercentageByQuality(validityPercentage, oField, page);
                }
                RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "Field: " + labelName + " 's validityPercentage is: " + validityPercentage);
                if (validityPercentage >= 90) // If validity is less than 90, set value to 1 otherwise keep it at 0
                {
                    oField.Status = 0; // If this is changed to 1, then the box will highlight in ICN, need to look a the validity percentage and agree what level to set this at
                }
                else
                {
                    oField.Status = 1;
                }
                RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "updateEntityAndProperties: status: " + oField.Status);
                if (altType != null)
                {
                    oField.Type = altType;
                }
                else
                {
                    oField.Type = labelName;
                }
                oField.set_Variable("entityType", typeName);
                oField.set_Variable("confidence", validityPercentage.ToString());

                oField.Text = text;
                RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "updateEntityAndProperties: Type: " + oField.Type + ", Text: " + text);
                if (location != null)
                {
                    oField.set_Variable("Position", location.getPosition());
                    RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "updateEntityAndProperties: Position: " + location.getPosition());
                    oField.set_Variable("l", location.getLeft() + "");
                    oField.set_Variable("t", location.getTop() + "");
                    oField.set_Variable("r", location.getRight() + "");
                    oField.set_Variable("b", location.getBottom() + "");
                }
                oField.set_Variable("validityPercentage", validityPercentage + "");
                RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "updateEntityAndProperties: Validity Percentage: " + validityPercentage);
                oField.set_Variable("label", labelName);
                oField.set_Variable("subMatch1", text);
                oField.set_Variable("entityName", name);

                if (extraData != null) // If extra data has values, add them to the field
                {
                    foreach (String key in extraData.Keys)
                    {
                        oField.set_Variable(key, extraData[key]);
                    }
                }

                if (similarity > 0.0)
                {
                    RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "updateEntityAndProperties: key similarity: " + similarity.ToString());
                    oField.set_Variable("KeySimilarity", similarity.ToString());
                }


            }

            RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "***End Utils: updateEntityAndProperties");
        }

        private int decreaseValidityPercentageByQuality(int validityPercentage, TDCOLib.IDCO oField, FieldObject page)
        {
            RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "Input validityPercentage: " + validityPercentage);
            int quality = calculateFieldQuality(oField, page);
            int decrease;
            RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "Field: " + oField.ID + " 's Quality is: " + quality);
            if (quality >= 90 || quality == 0)
                decrease = 0;
            else if (quality >= 70 && quality < 90)
                decrease = 5;
            else if (quality >= 50 && quality < 70)
                decrease = 7;
            else
                decrease = 11;
            return validityPercentage - decrease;
        }

        public int calculateFieldQuality(TDCOLib.IDCO oField, FieldObject page)
        {
            int pLeft, pTop, pRight, pBottom;
            getDCOLocation(oField, out pLeft, out pTop, out pRight, out pBottom);
            RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "oField location: " + pLeft + ", " + pTop + ", " + pRight + ", " + pBottom);
            PageLocation location = new PageLocation(Int32.Parse(pLeft.ToString()), Int32.Parse(pTop.ToString()), Int32.Parse(pRight.ToString()), Int32.Parse(pBottom.ToString()));
            FieldObject fo = page.getItemForLocation(location);
            int result = 0;
            if (fo != null)
            {
                result = fo.getConfidence();
                RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "quality for FO: " + fo.getAllLines() + " is: " + result);
            }
            return result;
        }

        public Boolean getDCOLocation(TDCOLib.IDCO field, out int left, out int top, out int right, out int bottom)
        {
            left = 0;
            top = 0;
            right = 0;
            bottom = 0;

            string[] fieldPosition = field.Variable["Position"].Split(',');

            if (fieldPosition.Length == 4)
            {
                int.TryParse(fieldPosition[0], out left);
                int.TryParse(fieldPosition[1], out top);
                int.TryParse(fieldPosition[2], out right);
                int.TryParse(fieldPosition[3], out bottom);
                return true;
            }
            else
            {
                return false;
            }

        }

    }
}
