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

namespace NPOI.HSLF.UserModel
{
    using Record;
    using System;
    using System.Collections.Generic;
    using Model;

    /**
     * Title masters define the design template for slides with a Title Slide layout.
     */
    public class HSLFTitleMaster : HSLFMasterSheet
    {
        private List<List<HSLFTextParagraph>> _paragraphs = new List<List<HSLFTextParagraph>>();

        /**
     * Constructs a TitleMaster
     *
     */
        public HSLFTitleMaster(Slide record, int sheetNo) : base(record, sheetNo)
        {
            foreach (List<HSLFTextParagraph> l in HSLFTextParagraph.findTextParagraphs(GetPPDrawing(), this))
            {
                if (!_paragraphs.Contains(l))
                {
                    _paragraphs.Add(l);
                }
            }
        }

        /**
     * Returns an array of all the TextRuns found
     */
        public override List<List<HSLFTextParagraph>> GetTextParagraphs()
        {
            return _paragraphs;
        }

        /**
     * Delegate the call to the underlying slide master.
     */
        public override TextPropCollection GetPropCollection(int txtype, int level, String name, bool isCharacter)
        {
            HSLFMasterSheet master = GetMasterSheet();
            return (master == null) ? null : master.GetPropCollection(txtype, level, name, isCharacter);
        }

        /**
     * Returns the slide master for this title master.
     */
        public override HSLFMasterSheet GetMasterSheet()
        {
            SlideAtom sa = ((Slide)GetSheetContainer()).getSlideAtom();
            int masterId = sa.getMasterID();
            foreach (HSLFSlideMaster sm in GetSlideShow().GetSlideMasters())
            {
                if (masterId == sm._getSheetNumber())
                {
                    return sm;
                }
            }

            return null;
        }
    }
}
