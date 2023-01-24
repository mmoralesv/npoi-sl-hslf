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
using NPOI.DDF;
using NPOI.HSLF.Exceptions;
using NPOI.Util;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace NPOI.HSLF.Record
{

    /**
     * An atom record that specifies whether a shape is a placeholder shape.
     * The number, position, and type of placeholder shapes are determined by
     * the slide layout as specified in the SlideAtom record.
     *
     * @since POI 3.14-Beta2
     */
    public class HSLFEscherClientDataRecord : EscherClientDataRecord {

        private List<Record> _childRecords = new List<Record>();

        public HSLFEscherClientDataRecord() { }

        public HSLFEscherClientDataRecord(HSLFEscherClientDataRecord other) {
            // TODO: for now only reference others children, later copy them when Record.copy is available
            // other._childRecords.stream().map(Record::copy).forEach(_childRecords::add);
            _childRecords.AddRange(other._childRecords);
        }


        public List<Record> getHSLFChildRecords() {
            return _childRecords;
        }

        public void removeChild(Record childClass) {
            _childRecords.Remove(childClass);
        }

        public void addChild(Record childRecord) {
            _childRecords.Add(childRecord);
        }

        public override int FillFields(byte[] data, int offset, IEscherRecordFactory recordFactory) {
            int bytesRemaining = ReadHeader(data, offset);
            byte[] remainingData = IOUtils.SafelyClone(data, offset+8, bytesRemaining, RecordAtom.GetMaxRecordLength());
            setRemainingData(remainingData);
            return bytesRemaining + 8;
        }

        public override int Serialize(int offset, byte[] data, EscherSerializationListener listener) {
            listener.BeforeRecordSerialize(offset, RecordId, this);

            LittleEndian.PutShort(data, offset, Options);
            LittleEndian.PutShort(data, offset+2, RecordId);

            byte[] childBytes = getRemainingData();

            LittleEndian.PutInt(data, offset+4, childBytes.Length);
            childBytes.CopyTo(data, offset+8);//        System.arraycopy(childBytes, 0, data, offset+8, childBytes.Length);
            int recordSize = 8+childBytes.Length;
            listener.AfterRecordSerialize(offset+recordSize, RecordId, recordSize, this);
            return recordSize;
        }

        public override int RecordSize
        {
            get { return 8 + RemainingData.Length; }
        }

        public override byte[] RemainingData { get { return new byte[0]; } set { } }
        //{
            //get {
            //    try {
            //        (OutputStream bos = new OutputStream())
            //        foreach (Record r in _childRecords) {
            //            r.WriteOut(bos);
            //        }
            //        return bos.toByteArray();
            //    } catch (IOException e) {
            //        throw new HSLFException(e);
            //    }
            //}
            //set {
            //    _childRecords.clear();
            //    int offset = 0;
            //    while (offset < remainingData.length) {
            //        org.apache.poi.hslf.record.Record r = Record.buildRecordAtOffset(remainingData, offset);
            //        if (r != null) {
            //            _childRecords.add(r);
            //        }
            //        long rlen = LittleEndian.getUInt(remainingData, offset+4);
            //        offset = Math.toIntExact(offset + 8 + rlen);
            //    }
            //} 
        //}

        public override string RecordName { get { return "HSLFClientData"; } }

        public HSLFEscherClientDataRecord copy() {
            return new HSLFEscherClientDataRecord(this);
        }
    }
}