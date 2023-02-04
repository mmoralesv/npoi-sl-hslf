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
using System.Text;

namespace NPOI.HSLF.Record
{

    /**
     * A Slide Atom (type 1007). Holds information on the parent Slide, what
     * Master Slide it uses, what Notes is attached to it, that sort of thing.
     * It also has a SSlideLayoutAtom embedded in it, but without the Atom header
     */

    public class SlideAtom : RecordAtom
    {
        public static int USES_MASTER_SLIDE_ID = unchecked((int)0x80000000);
        // private static int MASTER_SLIDE_ID      =  0x00000000;

        private byte[] _header;
        private static long _type = 1007L;

        private int masterID;
        private int notesID;

        private bool followMasterObjects;
        private bool followMasterScheme;
        private bool followMasterBackground;
        private SlideAtomLayout layoutAtom;
        private byte[] reserved;


        /** Get the ID of the master slide used. 0 if this is a master slide, otherwise -2147483648 */
        public int GetMasterID() { return masterID; }
        /** Change slide master.  */
        public void SetMasterID(int id) { masterID = id; }
        /** Get the ID of the notes for this slide. 0 if doesn't have one */
        public int GetNotesID() { return notesID; }
        /** Get the embedded SSlideLayoutAtom */
        public SlideAtomLayout GetSSlideLayoutAtom() { return layoutAtom; }

        /** Change the ID of the notes for this slide. 0 if it no longer has one */
        public void SetNotesID(int id) { notesID = id; }

        public bool GetFollowMasterObjects() { return followMasterObjects; }
        public bool GetFollowMasterScheme() { return followMasterScheme; }
        public bool GetFollowMasterBackground() { return followMasterBackground; }
        public void SetFollowMasterObjects(bool flag) { followMasterObjects = flag; }
        public void SetFollowMasterScheme(bool flag) { followMasterScheme = flag; }
        public void SetFollowMasterBackground(bool flag) { followMasterBackground = flag; }


        /* *************** record code follows ********************** */

        /**
         * For the Slide Atom
         */
        protected SlideAtom(byte[] source, int start, int len)
        {
            // Sanity Checking
            if (len < 30) { len = 30; }

            // Get the header
            _header = Arrays.CopyOfRange(source, start, start + 8);

            // Grab the 12 bytes that is "SSlideLayoutAtom"
            byte[] SSlideLayoutAtomData = Arrays.CopyOfRange(source, start + 8, start + 12 + 8);
            // Use them to build up the SSlideLayoutAtom
            layoutAtom = new SlideAtomLayout(SSlideLayoutAtomData);

            // Get the IDs of the master and notes
            masterID = LittleEndian.GetInt(source, start + 12 + 8);
            notesID = LittleEndian.GetInt(source, start + 16 + 8);

            // Grok the flags, stored as bits
            int flags = LittleEndian.GetUShort(source, start + 20 + 8);
            followMasterBackground = (flags & 4) == 4;
            followMasterScheme = (flags & 2) == 2;
            followMasterObjects = (flags & 1) == 1;

            // If there's any other bits of data, keep them about
            // 8 bytes header + 20 bytes to flags + 2 bytes flags = 30 bytes
            reserved = IOUtils.SafelyClone(source, start + 30, len - 30, GetMaxRecordLength());
        }

        /**
         * Create a new SlideAtom, to go with a new Slide
         */
        public SlideAtom()
        {
            _header = new byte[8];
            LittleEndian.PutUShort(_header, 0, 2);
            LittleEndian.PutUShort(_header, 2, (int)_type);
            LittleEndian.PutInt(_header, 4, 24);

            byte[] ssdate = new byte[12];
            layoutAtom = new SlideAtomLayout(ssdate);
            layoutAtom.SetGeometryType(SlideAtomLayout.SlideLayoutType.BLANK_SLIDE);

            followMasterObjects = true;
            followMasterScheme = true;
            followMasterBackground = true;
            masterID = USES_MASTER_SLIDE_ID; // -2147483648;
            notesID = 0;
            reserved = new byte[2];
        }

        /**
         * We are of type 1007
         */
        public override long GetRecordType() { return _type; }

        /**
         * Write the contents of the record back, so it can be written
         *  to disk
         */

        public override void WriteOut(OutputStream _out)
        {
            // Header
            _out.Write(_header);

            // SSSlideLayoutAtom stuff
            layoutAtom.WriteOut(_out);

            // IDs
            WriteLittleEndian(masterID, _out);
            WriteLittleEndian(notesID, _out);

            // Flags
            short flags = 0;
            if (followMasterObjects) { flags += (short)1; }
            if (followMasterScheme) { flags += (short)2; }
            if (followMasterBackground) { flags += (short)4; }
            WriteLittleEndian(flags, _out);

            // Reserved data
            _out.Write(reserved);
        }

        public override IDictionary<string, Func<object>> GetGenericProperties()
        {
            throw new NotImplementedException();
        }
        
        public IDictionary<string, Func<T>> GetGenericProperties<T>()
        {
            return (IDictionary<string, Func<T>>)GenericRecordUtil.GetGenericProperties(
                "masterID", () => GetMasterID(),
                "notesID", () => GetNotesID(),
                "followMasterObjects", () => GetFollowMasterObjects(),
                "followMasterScheme", () => GetFollowMasterScheme(),
                "followMasterBackground", () => GetFollowMasterBackground(),
                "layoutAtom", GetSSlideLayoutAtom
            );
        }
    }
}