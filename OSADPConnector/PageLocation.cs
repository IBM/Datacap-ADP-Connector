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
    class PageLocation
    {
        private int top;
        private int bottom;
        private int left;
        private int right;

        public PageLocation(int left, int top, int right, int bottom)
        {
            this.top = top;
            this.bottom = bottom;
            this.left = left;
            this.right = right;
        }

        public String getPosition()
        {
            return this.left + "," + this.top + "," + this.right + "," + this.bottom;
        }

        public void expand(PageLocation expandInto)
        {
            if (expandInto != null)
            {
                this.setBottom(Math.Max(this.getBottom(), expandInto.getBottom()));
                this.setTop(Math.Min(this.getTop(), expandInto.getTop()));
                this.setLeft(Math.Min(this.getLeft(), expandInto.getLeft()));
                this.setRight(Math.Max(this.getRight(), expandInto.getRight()));
            }
        }
        public void setTop(int top)
        {
            this.top = top;
        }
        public int getTop()
        {
            return this.top;
        }

        public void setBottom(int bottom)
        {
            this.bottom = bottom;
        }
        public int getBottom()
        {
            return this.bottom;
        }

        public void setLeft(int left)
        {
            this.left = left;
        }
        public int getLeft()
        {
            return this.left;
        }

        public void setRight(int right)
        {
            this.right = right;
        }
        public int getRight()
        {
            return this.right;
        }

        public Boolean isEquivalent(PageLocation that)
        {
            bool result;
            if (that == null)
            {
                result = false;
            }
            else
            {
                result = this.getLeft() == that.getLeft() && this.getTop() == that.getTop() && this.getRight() == that.getRight() && this.getBottom() == that.getBottom();
            }
            return result;
        }

    }
}
