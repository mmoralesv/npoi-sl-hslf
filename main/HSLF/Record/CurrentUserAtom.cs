
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

using NPOI.HSLF.Exceptions;
using NPOI.POIFS.FileSystem;
using NPOI.Util;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace NPOI.HSLF.Record
{
    /**
     * This is a special kind of Atom, because it doesn't live inside the
     *  PowerPoint document. Instead, it lives in a separate stream in the
     *  document. As such, it has to be treated specially
     */
    public class CurrentUserAtom
    {
        //private static Logger LOG = LogManager.getLogger(CurrentUserAtom.class);

        /** Standard Atom header */
        private static byte[] atomHeader = new byte[] { 0, 0, -10, 15 };
        /** The PowerPoint magic number for a non-encrypted file */
        private static byte[] headerToken = new byte[] { 95, -64, -111, -29 };
        /** The PowerPoint magic number for an encrypted file */
        private static byte[] encHeaderToken = new byte[] { -33, -60, -47, -13 };
        // The Powerpoint 97 version, major and minor numbers
        // byte[] ppt97FileVer = new byte[] { 8, 00, -13, 03, 03, 00 };

        /** The version, major and minor numbers */
        private int docFinalVersion;
        private byte docMajorNo;
        private byte docMinorNo;

        /** The Offset into the file for the current edit */
        private long currentEditOffset;
        /** The Username of the last person to edit the file */
        private String lastEditUser;
        /** The document release version. Almost always 8 */
        private long releaseVersion;

        /** Only correct after reading in or writing out */
        private byte[] _contents;

        /** Flag for encryption state of the whole file */
        private bool isEncrypted;


        /* ********************* getter/setter follows *********************** */

        public int getDocFinalVersion() { return docFinalVersion; }
        public byte getDocMajorNo() { return docMajorNo; }
        public byte getDocMinorNo() { return docMinorNo; }

        public long getReleaseVersion() { return releaseVersion; }
        public void setReleaseVersion(long rv) { releaseVersion = rv; }

        /** Points to the UserEditAtom */
        public long getCurrentEditOffset() { return currentEditOffset; }
        public void setCurrentEditOffset(long id) { currentEditOffset = id; }

        public String getLastEditUsername() { return lastEditUser; }
        public void setLastEditUsername(String u) { lastEditUser = u; }

        public bool IsEncrypted() { return isEncrypted; }
        public void SetEncrypted(bool isEncrypted) { this.isEncrypted = isEncrypted; }


        /* ********************* real code follows *************************** */

        /**
         * Create a new Current User Atom
         */
        public CurrentUserAtom()
        {
            _contents = new byte[0];

            // Initialise to empty
            docFinalVersion = 0x03f4;
            docMajorNo = 3;
            docMinorNo = 0;
            releaseVersion = 8;
            currentEditOffset = 0;
            lastEditUser = "Apache POI";
            isEncrypted = false;
        }


        /**
         * Find the Current User in the filesystem, and create from that
         */
        public CurrentUserAtom(DirectoryNode dir)
        {
            // Decide how big it is
            DocumentEntry docProps =
                (DocumentEntry)dir.GetEntry("Current User");

            // If it's clearly junk, bail out
            if (docProps.Size > 131072)
            {
                throw new CorruptPowerPointFileException("The Current User stream is implausably long. It's normally 28-200 bytes long, but was " + docProps.getSize() + " bytes");
            }

            // Grab the contents
            using (InputStream _in = dir.CreateDocumentInputStream("Current User"))
            {
                _contents = IOUtils.ToByteArray(_in, docProps.Size, getMaxRecordLength());
            }

            // See how long it is. If it's under 28 bytes long, we can't
            //  read it
            if (_contents.Length < 28)
            {
                bool isPP95 = dir.HasEntry(PP95_DOCUMENT);
                // PPT95 has 4 byte size, then data
                if (!isPP95 && _contents.Length >= 4)
                {
                    int size = LittleEndian.GetInt(_contents);
                    isPP95 = (size + 4 == _contents.Length);
                }

                if (isPP95)
                {
                    throw new OldPowerPointFormatException("Based on the Current User stream, you seem to have supplied a PowerPoint95 file, which isn't supported");
                }
                else
                {
                    throw new CorruptPowerPointFileException("The Current User stream must be at least 28 bytes long, but was only " + _contents.length);
                }
            }

            // Set everything up
            init();
        }

        /**
         * Actually do the creation from a block of bytes
         */
        private void init()
        {
            // First up is the size, in 4 bytes, which is fixed
            // Then is the header

            isEncrypted = (LittleEndian.GetInt(encHeaderToken) == LittleEndian.GetInt(_contents, 12));

            // Grab the edit offset
            currentEditOffset = LittleEndian.GetUInt(_contents, 16);

            // Grab the versions
            docFinalVersion = LittleEndian.GetUShort(_contents, 22);
            docMajorNo = _contents[24];
            docMinorNo = _contents[25];

            // Get the username length
            long usernameLen = LittleEndian.GetUShort(_contents, 20);
            if (usernameLen > 512)
            {
                // Handle the case of it being garbage
                //LOG.atWarn().log("Invalid username length {} found, treating as if there was no username set", box(usernameLen));
                usernameLen = 0;
            }

            // Now we know the length of the username,
            //  use this to grab the revision
            if (_contents.Length >= 28 + (int)usernameLen + 4)
            {
                releaseVersion = LittleEndian.GetUInt(_contents, 28 + (int)usernameLen);
            }
            else
            {
                // No revision given, as not enough data. Odd
                releaseVersion = 0;
            }

            // Grab the unicode username, if stored
            int start = 28 + (int)usernameLen + 4;

            if (_contents.Length >= start + 2 * usernameLen)
            {
                lastEditUser = StringUtil.GetFromUnicodeLE(_contents, start, (int)usernameLen);
            }
            else
            {
                // Fake from the 8 bit version
                lastEditUser = StringUtil.GetFromCompressedUnicode(_contents, 28, (int)usernameLen);
            }
        }


        /**
         * Writes ourselves back out
         */
        public void WriteOut(OutputStream _out)
        {
            // Decide on the size
            //  8 = atom header
            //  20 = up to name
            //  4 = revision
            //  3 * len = ascii + unicode
            int size = 8 + 20 + 4 + (3 * lastEditUser.Length);
            _contents = IOUtils.SafelyAllocate(size, getMaxRecordLength());

            // First we have a 8 byte atom header
            Array.Copy(atomHeader, 0, _contents, 0, 4);
            // Size is 20+user len + revision len(4)
            int atomSize = 20 + 4 + lastEditUser.Length;
            LittleEndian.PutInt(_contents, 4, atomSize);

            // Now we have the size of the details, which is 20
            LittleEndian.PutInt(_contents, 8, 20);

            // Now the ppt un-encrypted header token (4 bytes)
            Array.Copy((isEncrypted ? encHeaderToken : headerToken), 0, _contents, 12, 4);

            // Now the current edit offset
            LittleEndian.PutInt(_contents, 16, (int)currentEditOffset);

            // The username gets stored twice, once as US
            //  ascii, and again as unicode laster on
            byte[] asciiUN = IOUtils.SafelyAllocate(lastEditUser.Length, getMaxRecordLength());
            StringUtil.PutCompressedUnicode(lastEditUser, asciiUN, 0);

            // Now we're able to do the length of the last edited user
            LittleEndian.PutShort(_contents, 20, (short)asciiUN.Length);

            // Now the file versions, 2+1+1
            LittleEndian.PutShort(_contents, 22, (short)docFinalVersion);
            _contents[24] = docMajorNo;
            _contents[25] = docMinorNo;

            // 2 bytes blank
            _contents[26] = 0;
            _contents[27] = 0;

            // At this point we have the username as us ascii
            Array.Copy(asciiUN, 0, _contents, 28, asciiUN.Length);

            // 4 byte release version
            LittleEndian.PutInt(_contents, 28 + asciiUN.Length, (int)releaseVersion);

            // username in unicode
            byte[] ucUN = IOUtils.SafelyAllocate(lastEditUser.Length * 2L, getMaxRecordLength());
            StringUtil.PutUnicodeLE(lastEditUser, ucUN, 0);
            Array.Copy(ucUN, 0, _contents, 28 + asciiUN.Length + 4, ucUN.Length);

            // Write out
            _out.Write(_contents);
        }

        /**
         * Writes ourselves back out to a filesystem
         */
        //public void WriteToFS(POIFSFileSystem fs) {
        //    // Grab contents
        //    using (MemoryStream baos = new MemoryStream()) {
        //        WriteOut((OutputStream)baos);
        //        using (InputStream _is = baos) {
        //            // Write out
        //            fs.CreateOrUpdateDocument(_is, "Current User");
        //        }
        //    }
        //}
    }
}
