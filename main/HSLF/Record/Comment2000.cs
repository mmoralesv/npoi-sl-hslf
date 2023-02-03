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

using NPOI.Util;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace NPOI.HSLF.Record
{
    /**
     * This class represents a comment on a slide, in the format used by
     *  PPT 2000/XP/etc. (PPT 97 uses plain Escher Text Boxes for comments)
     */
    public class Comment2000 : RecordContainer
    {
        private byte[] _header;
        private static long _type = RecordTypes.Comment2000.typeID;

        // Links to our more interesting children

        /**
     * An optional string that specifies the name of the author of the presentation comment.
     */
        private CString authorRecord;

        /**
     * An optional string record that specifies the text of the presentation comment
     */
        private CString authorInitialsRecord;

        /**
     * An optional string record that specifies the initials of the author of the presentation comment
     */
        private CString commentRecord;

        /**
     * A Comment2000Atom record that specifies the settings for displaying the presentation comment
     */
        private Comment2000Atom commentAtom;

        /**
     * Returns the Comment2000Atom of this Comment
     */
        public Comment2000Atom GetComment2000Atom()
        {
            return commentAtom;
        }

        /**
     * Get the Author of this comment
     */
        public String GetAuthor()
        {
            return authorRecord == null ? null : authorRecord.GetText();
        }

        /**
     * Set the Author of this comment
     */
        public void SetAuthor(String author)
        {
            authorRecord.SetText(author);
        }

        /**
     * Get the Author's Initials of this comment
     */
        public String GetAuthorInitials()
        {
            return authorInitialsRecord == null ? null : authorInitialsRecord.GetText();
        }

        /**
     * Set the Author's Initials of this comment
     */
        public void SetAuthorInitials(String initials)
        {
            authorInitialsRecord.SetText(initials);
        }

        /**
     * Get the text of this comment
     */
        public String GetText()
        {
            return commentRecord == null ? null : commentRecord.GetText();
        }

        /**
     * Set the text of this comment
     */
        public void SetText(String text)
        {
            commentRecord.SetText(text);
        }

        /**
     * Set things up, and find our more interesting children
     */
        protected Comment2000(byte[] source, int start, int len)
        {
            // Grab the header
            int elementsToTake = 8;
            List<byte> tmpNewArray = new List<byte>(elementsToTake);
            tmpNewArray.AddRange(source.Skip(start).Take(elementsToTake));
            if (tmpNewArray.Count < elementsToTake)
            {
                int emptySpace = elementsToTake - tmpNewArray.Count;
                tmpNewArray.InsertRange(tmpNewArray.Count - 1, Enumerable.Repeat(byte.MinValue, emptySpace));
            }

            _header = tmpNewArray.ToArray();

            // Find our children
            _children = FindChildRecords(source, start + 8, len - 8);
            FindInterestingChildren();
        }

        /**
     * Go through our child records, picking out the ones that are
     *  interesting, and saving those for use by the easy helper
     *  methods.
     */
        private void FindInterestingChildren()
        {
            foreach (Record r in _children)
            {
                if (r is CString)
                {
                    CString cs = (CString)r;
                    int recInstance = cs.GetOptions() >> 4;
                    switch (recInstance)
                    {
                        case 0:
                            authorRecord = cs;
                            break;
                        case 1:
                            commentRecord = cs;
                            break;
                        case 2:
                            authorInitialsRecord = cs;
                            break;
                        default: break;
                    }
                }
                else if (r is Comment2000Atom)
                {
                    commentAtom = (Comment2000Atom)r;
                }
                else
                {
                    //LOG.atWarn().log("Unexpected record with type={} in Comment2000: {}", box(r.getRecordType()),r.getClass().getName());
                }
            }
        }

        /**
     * Create a new Comment2000, with blank fields
     */
        public Comment2000()
        {
            _header = new byte[8];
            _children = new Record[4];

            // Setup our header block
            _header[0] = 0x0f; // We are a container record
            LittleEndian.PutShort(_header, 2, (short)_type);

            // Setup our child records
            CString csa = new CString();
            CString csb = new CString();
            CString csc = new CString();
            csa.SetOptions(0x00);
            csb.SetOptions(0x10);
            csc.SetOptions(0x20);
            _children[0] = csa;
            _children[1] = csb;
            _children[2] = csc;
            _children[3] = new Comment2000Atom();
            FindInterestingChildren();
        }

        /**
     * We are of type 1200
     */
        public override long GetRecordType()
        {
            return _type;
        }

        /**
     * Write the contents of the record back, so it can be written
     *  to disk
     */
        public override void WriteOut(OutputStream outputStream)
        {
            try
            {
                WriteOut(_header[0], _header[1], _type, _children, outputStream);
            }
            catch (Exception e)
            {
                throw new IOException();
            }
        }
    }
}