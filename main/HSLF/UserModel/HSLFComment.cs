/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
   The ASF licenses this file to You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */

using NPOI.Common.UserModel;
using NPOI.HSLF.Record;
using NPOI.SL.UserModel;
using NPOI.Util;
using System;
using System.Collections.Generic;

namespace NPOI.HSLF.UserModel
{
    public class HSLFComment : Comment
    {
        private Comment2000 _comment2000;

        public HSLFComment(Comment2000 comment2000)
        {
            _comment2000 = comment2000;
        }

        protected Comment2000 GetComment2000()
        {
            return _comment2000;
        }

        /**
     * Get the Author of this comment
     */
        public string GetAuthor()
        {
            return _comment2000.GetAuthor();
        }

        /**
     * Set the Author of this comment
     */
        public void SetAuthor(String author)
        {
            _comment2000.SetAuthor(author);
        }

        /**
     * Get the Author's Initials of this comment
     */
        public String GetAuthorInitials()
        {
            return _comment2000.GetAuthorInitials();
        }

        /**
     * Set the Author's Initials of this comment
     */
        public void SetAuthorInitials(String initials)
        {
            _comment2000.SetAuthorInitials(initials);
        }

        /**
     * Get the text of this comment
     */
        public String GetText()
        {
            return _comment2000.GetText();
        }

        /**
     * Set the text of this comment
     */
        public void SetText(String text)
        {
            _comment2000.SetText(text);
        }

        DateTime Comment.GetDate()
        {
            throw new NotImplementedException();
        }

        public DateTime GetDate()
        {
            return _comment2000.GetComment2000Atom().GetDate();
        }

        public void SetDate(DateTime date)
        {
            _comment2000.GetComment2000Atom().SetDate(date);
        }

        public Point2D GetOffset()
        {
            double x = Units.MasterToPoints(_comment2000.GetComment2000Atom().GetXOffset());
            double y = Units.MasterToPoints(_comment2000.GetComment2000Atom().GetYOffset());
            return new Point2D.Double(x, y);
        }

        public void SetOffset(Point2D offset)
        {
            // int x = Units.PointsToMaster(offset.GetX());
            // int y = Units.PointsToMaster(offset.GetY());
            // _comment2000.GetComment2000Atom().SetXOffset(x);
            // _comment2000.GetComment2000Atom().SetYOffset(y);
        }
    }
}