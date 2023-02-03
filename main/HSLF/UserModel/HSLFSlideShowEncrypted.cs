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
using NPOI.HSLF.Exceptions;
using NPOI.HSLF.Record;
using NPOI.HSSF.Record.Crypto;
using NPOI.POIFS.Crypt;
using NPOI.POIFS.FileSystem;
using NPOI.SL.UserModel;
using NPOI.Util;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace NPOI.HSLF.UserModel
{

    /**
     * This class provides helper functions for encrypted PowerPoint documents.
     */

    public class HSLFSlideShowEncrypted : ICloseable
    {
        DocumentEncryptionAtom dea;
        EncryptionInfo _encryptionInfo;
        //    Cipher cipher = null;
        ChunkedCipherOutputStream cyos;

        private static BitField fieldRecInst = new BitField(0xFFF0);

        private static int[] BLIB_STORE_ENTRY_PARTS = {
            1,     // btWin32
            1,     // btMacOS
            16,    // rgbUid
            2,     // tag
            4,     // size
            4,     // cRef
            4,     // foDelay
            1,     // unused1
            1,     // cbName (@ index 33)
            1,     // unused2
            1,     // unused3
    };

        public HSLFSlideShowEncrypted(DocumentEncryptionAtom dea)
        {
            this.dea = dea;
        }

		public HSLFSlideShowEncrypted(byte[] docstream, Dictionary<int, Record.Record> recordMap)
        {
            // check for DocumentEncryptionAtom, which would be at the last offset
            // need to ignore already set UserEdit and PersistAtoms
            UserEditAtom userEditAtomWithEncryption = null;
            foreach (var me in recordMap)
            {
                Record.Record rec = me.Value;
                if (!(rec is UserEditAtom))
                {
                    continue;
                }
                UserEditAtom uea = (UserEditAtom)rec;
                if (uea.getEncryptSessionPersistIdRef() != -1)
                {
                    userEditAtomWithEncryption = uea;
                    break;
                }
            }

            if (userEditAtomWithEncryption == null)
            {
                dea = null;
                return;
            }

            Record.Record r = recordMap[userEditAtomWithEncryption.getPersistPointersOffset()];
            if (!(r is PersistPtrHolder))
            {
                throw new RecordFormatException("Encountered an unexpected record-type: " + r);
            }
            PersistPtrHolder ptr = (PersistPtrHolder)r;

            int encOffset = ptr.getSlideLocationsLookup()[userEditAtomWithEncryption.getEncryptSessionPersistIdRef()];
            if (encOffset == null)
            {
                // encryption info doesn't exist anymore
                // SoftMaker Freeoffice produces such invalid files - check for "SMNativeObjData" ole stream
                dea = null;
                return;
            }

            r = recordMap[encOffset];
            if (r == null)
            {
                r = Record.Record.BuildRecordAtOffset(docstream, encOffset);
                recordMap.Add(encOffset, r);
            }

            if (!(r is DocumentEncryptionAtom))
            {
                throw new EncryptedPowerPointFileException("Did not have a DocumentEncryptionAtom: " + r);
            }
            this.dea = (DocumentEncryptionAtom)r;

            String pass = Biff8EncryptionKey.CurrentUserPassword;
            EncryptionInfo ei = GetEncryptionInfo();
            try
            {
                if (ei == null || ei.Decryptor == null)
                {
                    throw new InvalidOperationException("Invalid encryption-info: " + ei);
                }
                if (!ei.Decryptor.VerifyPassword(pass != null ? pass : Decryptor.DEFAULT_PASSWORD))
                {
                    throw new EncryptedPowerPointFileException("PowerPoint file is encrypted. The correct password needs to be set via Biff8EncryptionKey.setCurrentUserPassword()");
                }
            }
            catch (Exception e)
            {
                throw new EncryptedPowerPointFileException(e);
            }
        }

        public DocumentEncryptionAtom GetDocumentEncryptionAtom()
        {
            return dea;
        }

        protected EncryptionInfo GetEncryptionInfo()
        {
            return (dea != null) ? dea.getEncryptionInfo() : null;
        }

        //protected OutputStream EncryptRecord(OutputStream plainStream, int persistId, Record.Record record)
        //{
        //    bool isPlain = (dea == null
        //        || record is UserEditAtom
        //        || record is PersistPtrHolder
        //        || record is DocumentEncryptionAtom
        //    );

        //    try
        //    {
        //        if (isPlain)
        //        {
        //            if (cyos != null)
        //            {
        //                // write cached data to stream
        //                cyos.Flush();
        //            }
        //            return plainStream;
        //        }

        //        if (cyos == null)
        //        {
        //            Encryptor enc = GetEncryptionInfo().Encryptor;
        //            //enc.SetChunkSize(-1);
        //            cyos = enc.GetDataStream(plainStream, 0);
        //        }
        //        cyos.InitCipherForBlock(persistId, false);
        //    }
        //    catch (Exception e)
        //    {
        //        throw new EncryptedPowerPointFileException(e);
        //    }
        //    return cyos;
        //}

        private static void readFully(ChunkedCipherInputStream ccis, byte[] docstream, int offset, int len)
        {
            if (IOUtils.ReadFully(ccis, docstream, offset, len) == -1)
            {
                throw new EncryptedPowerPointFileException("unexpected EOF");
            }
        }

        //protected void DecryptRecord(byte[] docstream, int persistId, int offset)
        //{
        //    if (dea == null)
        //    {
        //        return;
        //    }

        //    Decryptor dec = GetEncryptionInfo().Decryptor;
        //    //dec.SetChunkSize(-1);
        //    try
        //    {
        //        using (LittleEndianByteArrayInputStream lei = new LittleEndianByteArrayInputStream(docstream, offset))
        //        {
        //            using (ChunkedCipherInputStream ccis = (ChunkedCipherInputStream)dec.GetDataStream(lei, docstream.Length - offset, 0))
        //            {
        //                ccis.InitCipherForBlock(persistId);

        //                // decrypt header and read length to be decrypted
        //                readFully(ccis, docstream, offset, 8);
        //                // decrypt the rest of the record
        //                int rlen = (int)LittleEndian.GetUInt(docstream, offset + 4);
        //                readFully(ccis, docstream, offset + 8, rlen);
        //            }
        //        }
        //    }
        //    catch (Exception e)
        //    {
        //        throw new EncryptedPowerPointFileException(e);
        //    }
        //}

        //private void decryptPicBytes(byte[] pictstream, int offset, int len)
        //{
        //    // when reading the picture elements, each time a segment is read, the cipher needs
        //    // to be reset (usually done when calling Cipher.doFinal)
        //    LittleEndianByteArrayInputStream lei = new LittleEndianByteArrayInputStream(pictstream, offset);
        //    Decryptor dec = GetEncryptionInfo().Decryptor;
        //    ChunkedCipherInputStream ccis = (ChunkedCipherInputStream)dec.GetDataStream(lei, len, 0);
        //    readFully(ccis, pictstream, offset, len);
        //    ccis.Close();
        //    //lei.Close();
        //}

        //protected void decryptPicture(byte[] pictstream, int offset)
        //{
        //    if (dea == null)
        //    {
        //        return;
        //    }

        //    try
        //    {
        //        // decrypt header and read length to be decrypted
        //        decryptPicBytes(pictstream, offset, 8);
        //        int recInst = fieldRecInst.GetValue(LittleEndian.GetUShort(pictstream, offset));
        //        int recType = LittleEndian.GetUShort(pictstream, offset + 2);
        //        int rlen = (int)LittleEndian.GetUInt(pictstream, offset + 4);
        //        offset += 8;
        //        int endOffset = offset + rlen;

        //        if (recType == 0xF007)
        //        {
        //            // TOOD: get a real example file ... to actual test the FBSE entry
        //            // not sure where the foDelay block is

        //            // File BLIP Store Entry (FBSE)
        //            foreach (int part in BLIB_STORE_ENTRY_PARTS)
        //            {
        //                decryptPicBytes(pictstream, offset, part);
        //            }
        //            offset += 36;

        //            int cbName = LittleEndian.GetUShort(pictstream, offset - 3);
        //            if (cbName > 0)
        //            {
        //                // read nameData
        //                decryptPicBytes(pictstream, offset, cbName);
        //                offset += cbName;
        //            }

        //            if (offset == endOffset)
        //            {
        //                return; // no embedded blip
        //            }
        //            // fall through, read embedded blip now

        //            // update header data
        //            decryptPicBytes(pictstream, offset, 8);
        //            recInst = fieldRecInst.GetValue(LittleEndian.GetUShort(pictstream, offset));
        //            recType = LittleEndian.GetUShort(pictstream, offset + 2);
        //            // rlen = (int)LittleEndian.getUInt(pictstream, offset+4);
        //            offset += 8;
        //        }

        //        int rgbUidCnt = (recInst == 0x217 || recInst == 0x3D5 || recInst == 0x46B || recInst == 0x543 ||
        //            recInst == 0x6E1 || recInst == 0x6E3 || recInst == 0x6E5 || recInst == 0x7A9) ? 2 : 1;

        //        // rgbUid 1/2
        //        for (int i = 0; i < rgbUidCnt; i++)
        //        {
        //            decryptPicBytes(pictstream, offset, 16);
        //            offset += 16;
        //        }

        //        int nextBytes;
        //        if (recType == 0xF01A || recType == 0XF01B || recType == 0XF01C)
        //        {
        //            // metafileHeader
        //            nextBytes = 34;
        //        }
        //        else
        //        {
        //            // tag
        //            nextBytes = 1;
        //        }

        //        decryptPicBytes(pictstream, offset, nextBytes);
        //        offset += nextBytes;

        //        int blipLen = endOffset - offset;
        //        decryptPicBytes(pictstream, offset, blipLen);
        //    }
        //    catch (Exception e)
        //    {
        //        throw new CorruptPowerPointFileException(e);
        //    }
        //}

        //protected void encryptPicture(byte[] pictstream, int offset)
        //{
        //    if (dea == null)
        //    {
        //        return;
        //    }

        //    ChunkedCipherOutputStream ccos = null;

        //    try
        //    {
        //        using (LittleEndianByteArrayOutputStream los = new LittleEndianByteArrayOutputStream(pictstream, offset))
        //        {
        //            Encryptor enc = GetEncryptionInfo().Encryptor;
        //            //enc.setChunkSize(-1);
        //            ccos = enc.GetDataStream(los, 0);
        //            int recInst = fieldRecInst.GetValue(LittleEndian.GetUShort(pictstream, offset));
        //            int recType = LittleEndian.GetUShort(pictstream, offset + 2);
        //            int rlen = (int)LittleEndian.GetUInt(pictstream, offset + 4);

        //            ccos.Write(pictstream, offset, 8);
        //            ccos.Flush();
        //            offset += 8;
        //            int endOffset = offset + rlen;

        //            if (recType == 0xF007)
        //            {
        //                // TOOD: get a real example file ... to actual test the FBSE entry
        //                // not sure where the foDelay block is

        //                // File BLIP Store Entry (FBSE)
        //                int cbName = LittleEndian.GetUShort(pictstream, offset + 33);

        //                foreach (int part in BLIB_STORE_ENTRY_PARTS)
        //                {
        //                    ccos.Write(pictstream, offset, part);
        //                    ccos.Flush();
        //                    offset += part;
        //                }

        //                if (cbName > 0)
        //                {
        //                    ccos.Write(pictstream, offset, cbName);
        //                    ccos.Flush();
        //                    offset += cbName;
        //                }

        //                if (offset == endOffset)
        //                {
        //                    return; // no embedded blip
        //                }
        //                // fall through, read embedded blip now

        //                // update header data
        //                recInst = fieldRecInst.GetValue(LittleEndian.GetUShort(pictstream, offset));
        //                recType = LittleEndian.GetUShort(pictstream, offset + 2);
        //                ccos.Write(pictstream, offset, 8);
        //                ccos.Flush();
        //                offset += 8;
        //            }

        //            int rgbUidCnt = (recInst == 0x217 || recInst == 0x3D5 || recInst == 0x46B || recInst == 0x543 ||
        //                recInst == 0x6E1 || recInst == 0x6E3 || recInst == 0x6E5 || recInst == 0x7A9) ? 2 : 1;

        //            for (int i = 0; i < rgbUidCnt; i++)
        //            {
        //                ccos.Write(pictstream, offset, 16); // rgbUid 1/2
        //                ccos.Flush();
        //                offset += 16;
        //            }

        //            if (recType == 0xF01A || recType == 0XF01B || recType == 0XF01C)
        //            {
        //                ccos.Write(pictstream, offset, 34); // metafileHeader
        //                offset += 34;
        //                ccos.Flush();
        //            }
        //            else
        //            {
        //                ccos.Write(pictstream, offset, 1); // tag
        //                offset += 1;
        //                ccos.Flush();
        //            }

        //            int blipLen = endOffset - offset;
        //            ccos.Write(pictstream, offset, blipLen);
        //            ccos.Flush();
        //        }
        //    }
        //    catch (Exception e)
        //    {
        //        throw new EncryptedPowerPointFileException(e);
        //    }
        //    finally
        //    {
        //        IOUtils.CloseQuietly(ccos);
        //    }
        //}

        protected Record.Record[] updateEncryptionRecord(Record.Record[] records)
        {
            String password = Biff8EncryptionKey.CurrentUserPassword;
            if (password == null)
            {
                if (dea == null)
                {
                    // no password given, no encryption record exits -> done
                    return records;
                }
                else
                {
                    // need to remove password data
                    dea = null;
                    return removeEncryptionRecord(records);
                }
            }
            else
            {
                // create password record
                if (dea == null)
                {
                    dea = new DocumentEncryptionAtom();
                }
                EncryptionInfo ei = dea.getEncryptionInfo();
                byte[] salt = ei.Verifier.Salt;
                Encryptor enc = GetEncryptionInfo().Encryptor;
                if (salt == null)
                {
                    enc.ConfirmPassword(password);
                }
                else
                {
                    byte[] verifier = ei.Decryptor.GetVerifier();
                    enc.ConfirmPassword(password, null, null, verifier, salt, null);
                }

                // move EncryptionRecord to last slide position
                records = NormalizeRecords(records);
                return AddEncryptionRecord(records, dea);
            }
        }

        /**
         * remove duplicated UserEditAtoms and merge PersistPtrHolder.
         * Before this method is called, make sure that the offsets are correct,
         * i.e. call {@link HSLFSlideShowImpl#updateAndWriteDependantRecords(OutputStream, Map)}
         */
        protected static Record.Record[] NormalizeRecords(Record.Record[] records)
        {
            // http://msdn.microsoft.com/en-us/library/office/gg615594(v=office.14).aspx
            // repeated slideIds can be overwritten, i.e. ignored

            UserEditAtom uea = null;
            PersistPtrHolder pph = null;
            Dictionary<int, int> slideLocations = new Dictionary<int, int>();
            Dictionary<int, Record.Record> recordMap = new Dictionary<int, Record.Record>();
            List<int> obsoleteOffsets = new List<int>();
            int duplicatedCount = 0;
            foreach (Record.Record r in records)
            {
                if (r is PositionDependentRecord)
                {
                    PositionDependentRecord pdr = (PositionDependentRecord)r;
                    if (pdr is UserEditAtom)
                    {
                        uea = (UserEditAtom)pdr;
                        continue;
                    }

                    if (pdr is PersistPtrHolder)
                    {
                        if (pph != null)
                        {
                            duplicatedCount++;
                        }
                        pph = (PersistPtrHolder)pdr;
                        foreach (var me in pph.getSlideLocationsLookup())
                        {
                            int oldOffset = slideLocations.Add(me.Key, me.Value);
                            if (oldOffset != 0)
                            {
                                obsoleteOffsets.Add(oldOffset);
                            }
                        }
                        continue;
                    }

                    recordMap.Add(pdr.GetLastOnDiskOffset(), r);
                }
            }

            if (uea == null || pph == null || uea.getPersistPointersOffset() != pph.getLastOnDiskOffset())
            {
                throw new EncryptedDocumentException("UserEditAtom and PersistPtrHolder must exist and their offset need to match.");
            }

            recordMap.Add(pph.getLastOnDiskOffset(), pph);
            recordMap.Add(uea.getLastOnDiskOffset(), uea);

            if (duplicatedCount == 0 && obsoleteOffsets.Count == 0)
            {
                return records;
            }

            uea.setLastUserEditAtomOffset(0);
            pph.clear();
            foreach (var me in slideLocations)
            {
                pph.addSlideLookup(me.Key, me.Value);
            }

            foreach (int oldOffset in obsoleteOffsets)
            {
                recordMap.Remove(oldOffset);
            }

            return recordMap.Values.ToArray();
        }


        protected static Record.Record[] removeEncryptionRecord(Record.Record[] records)
        {
            int deaSlideId = -1;
            int deaOffset = -1;
            PersistPtrHolder ptr = null;
            UserEditAtom uea = null;
            List<Record.Record> recordList = new List<Record.Record>();
            foreach (Record.Record r in records)
            {
                if (r is DocumentEncryptionAtom)
                {
                    deaOffset = ((DocumentEncryptionAtom)r).getLastOnDiskOffset();
                    continue;
                }
                else if (r is UserEditAtom)
                {
                    uea = (UserEditAtom)r;
                    deaSlideId = uea.getEncryptSessionPersistIdRef();
                    uea.setEncryptSessionPersistIdRef(-1);
                }
                else if (r is PersistPtrHolder)
                {
                    ptr = (PersistPtrHolder)r;
                }
                recordList.Add(r);
            }

            if (ptr == null || uea == null)
            {
                throw new EncryptedDocumentException("UserEditAtom or PersistPtrholder not found.");
            }
            if (deaSlideId == -1 && deaOffset == -1)
            {
                return records;
            }

            Dictionary<int, int> tm = new Dictionary<int, int>(ptr.getSlideLocationsLookup());
            ptr.clear();
            int maxSlideId = -1;
            foreach (var me in tm)
            {
                if (me.Key == deaSlideId || me.Value == deaOffset)
                {
                    continue;
                }
                ptr.addSlideLookup(me.Key, me.Value);
                maxSlideId = Math.Max(me.Key, maxSlideId);
            }

            uea.setMaxPersistWritten(maxSlideId);

            records = recordList.ToArray();

            return records;
        }


        protected static Record.Record[] AddEncryptionRecord(Record.Record[] records, DocumentEncryptionAtom dea)
        {
            if (dea == null) { return null; };
            int ueaIdx = -1, ptrIdx = -1, deaIdx = -1, idx = -1;
            foreach (Record.Record r in records)
            {
                idx++;
                if (r is UserEditAtom)
                {
                    ueaIdx = idx;
                }
                else if (r is PersistPtrHolder)
                {
                    ptrIdx = idx;
                }
                else if (r is DocumentEncryptionAtom)
                {
                    deaIdx = idx;
                }
            }

            if (ueaIdx != -1 && ptrIdx != -1 && ptrIdx < ueaIdx)
            {
                if (deaIdx != -1)
                {
                    DocumentEncryptionAtom deaOld = (DocumentEncryptionAtom)records[deaIdx];
                    dea.setLastOnDiskOffset(deaOld.getLastOnDiskOffset());
                    records[deaIdx] = dea;
                    return records;
                }
                else
                {
                    PersistPtrHolder ptr = (PersistPtrHolder)records[ptrIdx];
                    UserEditAtom uea = ((UserEditAtom)records[ueaIdx]);
                    dea.setLastOnDiskOffset(ptr.getLastOnDiskOffset() - 1);
                    int nextSlideId = uea.getMaxPersistWritten() + 1;
                    ptr.addSlideLookup(nextSlideId, ptr.getLastOnDiskOffset() - 1);
                    uea.setEncryptSessionPersistIdRef(nextSlideId);
                    uea.setMaxPersistWritten(nextSlideId);

                    Record.Record[] newRecords = new Record.Record[records.Length + 1];
                    if (ptrIdx > 0)
                    {
                        Array.Copy(records, 0, newRecords, 0, ptrIdx);
                    }
                    if (ptrIdx < records.Length - 1)
                    {
                        Array.Copy(records, ptrIdx, newRecords, ptrIdx + 1, records.Length - ptrIdx);
                    }
                    newRecords[ptrIdx] = dea;
                    return newRecords;
                }
            }
            return null;
        }

        public void Close()
        {
            if (cyos != null)
            {
                cyos.Close();
            }
        }
    }
}
