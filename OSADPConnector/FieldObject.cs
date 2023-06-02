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
    class FieldObject
    {
        public const int FIELD_TYPE_NONE = 0;
        public const int FIELD_TYPE_BLOCK = 1;
        public const int FIELD_TYPE_PARAGRAPH = 2;
        public const int FIELD_TYPE_LINE = 3;
        public const int FIELD_TYPE_TABLE = 4;
        public const int FIELD_TYPE_PICTURE = 5;
        public const int FIELD_TYPE_CELL = 6;
        public const int FIELD_TYPE_ROW = 7;
        public const int FIELD_TYPE_PAGE = 8;
        public const int FIELD_TYPE_WORD = 9;
        public const int FIELD_TYPE_WORDGROUP = 10;
        public const int FIELD_TYPE_WORDGROUPLINE = 11;
        public const int FIELD_TYPE_WORDGROUPBLOCK = 12;
        public const int FIELD_TYPE_WGTABLE = 13;
        public const int FIELD_TYPE_WGROW = 14;
        public const int FIELD_TYPE_WGCOLUMN = 15;
        public const int FIELD_TYPE_ROWGROUP = 16;

        private static int KEYS_ONLY = 0;
        private static int VALUES_ONLY = 1;
        private static int KEYS_AND_VALUES = 2;

        private List<FieldObject> children = new List<FieldObject>();
        private Int32 confidence;
        private Boolean fieldIsKey = false;
        private Int32 fontSize;
        private List<String> lines = new List<String>();
        private List<String> originalLines = new List<String>();
        private PageLocation location = null;
        private List<PageLocation> locations = new List<PageLocation>();
        private List<Int32> lineQualityList = new List<Int32>();
        private List<String> lineFontSizeList = new List<String>();

        private int fieldType;
        private dclogXLib.IDCLog RRLog;
        private Utils utils = null;

        public FieldObject(int fieldType, dclogXLib.IDCLog RRLog)
        {
            this.fieldType = fieldType;
            this.RRLog = RRLog;
            if (this.RRLog == null)
            {
                throw new Exception("RRLog is null!!!");
            }
            this.utils = new Utils(this.RRLog);
        }

        public void add(FieldObject fo, Boolean adjustLocation = false)
        {
            children.Add((FieldObject)fo);
            if (adjustLocation)
            {
                this.location.expand(fo.getLocation());
            }
        }

        public void add(String line, PageLocation location)
        {
            lines.Add(line);
            locations.Add((PageLocation)location);
        }

        public void add(String line, PageLocation location, Int32 quality, String fontSize)
        {
            lines.Add(line);
            locations.Add((PageLocation)location);
            lineQualityList.Add(quality);
            lineFontSizeList.Add(fontSize);
        }

        public void setLines(List<String> lines)
        {
            this.lines = lines;
        }

        public void setOriginalLines(List<String> lines)
        {
            this.originalLines = lines;
        }

        public void setFontSize(int fontSize)
        {
            this.fontSize = fontSize;
        }

        public int getFieldType()
        {
            return fieldType;
        }

        public int getFontSize()
        {
            return this.fontSize;
        }

        public String getAllLines()
        {
            StringBuilder allLines = new StringBuilder("");

            for (int i = 0; i < lines.Count; i++)
            {
                allLines.Append(lines[i].Trim()).Append(" "); // SSM 2018-08-26
            }

            return allLines.ToString();
        }

        public FieldObject[] getChildren()
        {
            return this.children.ToArray();
        }

        public void setConfidence(Int32 confidence)
        {
            this.confidence = confidence;
        }
        public Int32 getConfidence()
        {
            return this.confidence;
        }

        public String[] getLines()
        {
            return this.lines.ToArray();
        }

        public void setLocation(PageLocation location)
        {
            this.location = location;
        }
        public PageLocation getLocation()
        {
            return this.location;
        }

        public Boolean isKey()
        {
            return this.fieldIsKey;
        }

        public FieldObject[] getChildren(int type)
        {
            return getChildren(type, KEYS_AND_VALUES);
        }

        private FieldObject[] getChildren(int type, int which)
        {
            List<FieldObject> entities = new List<FieldObject>();

            if (type == FIELD_TYPE_CELL)
            {
                foreach (FieldObject child in children)
                {
                    if (child.getFieldType() == FIELD_TYPE_TABLE)
                    {
                        FieldObject[] rows = child.getChildren();
                        for (int i = 0; i < rows.Length; i++)
                        {
                            if (rows[i].getFieldType() == FIELD_TYPE_ROW)
                            {
                                FieldObject[] cells = rows[i].getChildren();

                                for (int a = 0; a < cells.Length; a++)
                                {
                                    if (cells[a].getFieldType() == FIELD_TYPE_CELL)
                                    {
                                        if (((which == KEYS_ONLY) && cells[a].isKey()) ||
                                            ((which == VALUES_ONLY) && !cells[a].isKey()) ||
                                            (which == KEYS_AND_VALUES))
                                        {
                                            entities.Add((FieldObject)cells[a]);
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
            else
            {
                foreach (FieldObject child in children)
                {
                    if (child.getFieldType() == type)
                    {
                        if (((which == KEYS_ONLY) && child.isKey()) ||
                            ((which == VALUES_ONLY) && !child.isKey()) ||
                            (which == KEYS_AND_VALUES))
                        {
                            entities.Add(child);
                        }
                    }
                }
            }

            return entities.ToArray();
        }

        public FieldObject getItemForLocation(PageLocation loc)
        {
            FieldObject[] wordGroups = this.getChildren(FieldObject.FIELD_TYPE_WORDGROUP);
            foreach (FieldObject wg in wordGroups)
            {
                if (wg.getLocation().isEquivalent(loc))
                {
                    RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "Found wordgroup by location");
                    return wg;
                }
            }

            FieldObject[] wgls = this.getChildren(FieldObject.FIELD_TYPE_WORDGROUPLINE);
            foreach (FieldObject wgl in wgls)
            {
                if (wgl.getLocation().isEquivalent(loc))
                {
                    RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "Found wordgroupline by location");
                    return wgl;
                }
            }

            FieldObject[] wgbs = this.getChildren(FieldObject.FIELD_TYPE_WORDGROUPBLOCK);
            foreach (FieldObject wgb in wgbs)
            {
                if (wgb.getLocation().isEquivalent(loc))
                {
                    RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "Found wordgroupblock by location");
                    return wgb;
                }
            }

            FieldObject[] words = this.getChildren(FieldObject.FIELD_TYPE_WORD);
            foreach (FieldObject word in words)
            {
                if (word.getLocation().isEquivalent(loc))
                {
                    RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "Found word by location");
                    return word;
                }
            }

            RRLog.WriteEx(OSADPConnector.LOGGING_DEBUG, "Not Found anything by location");
            return null;
        }

    }
}
