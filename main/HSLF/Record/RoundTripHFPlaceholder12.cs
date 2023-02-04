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
     * An atom record that specifies that a shape is a header or footer placeholder shape
     */
    public class RoundTripHFPlaceholder12 : RecordAtom
    {
        /**
         * Record header.
         */
        private byte[] _header;

        /**
         * Specifies the placeholder shape ID.
         *
         * MUST be {@link OEPlaceholderAtom#MasterDate},  {@link OEPlaceholderAtom#MasterSlideNumber},
         * {@link OEPlaceholderAtom#MasterFooter}, or {@link OEPlaceholderAtom#MasterHeader}
         */
        private byte _placeholderId;

        /**
         * Create a new instance of <code>RoundTripHFPlaceholder12</code>
         */
        public RoundTripHFPlaceholder12()
        {
            _header = new byte[8];
            LittleEndian.PutUShort(_header, 0, 0);
            LittleEndian.PutUShort(_header, 2, (int)GetRecordType());
            LittleEndian.PutInt(_header, 4, 8);
            _placeholderId = 0;
        }

        /**
         * Constructs the comment atom record from its source data.
         *
         * @param source the source data as a byte array.
         * @param start the start offset into the byte array.
         * @param len the length of the slice in the byte array.
         */
        protected RoundTripHFPlaceholder12(byte[] source, int start, int len)
        {
            // Get the header.
            _header = Arrays.CopyOfRange(source, start, start + 8);

            // Get the record data.
            _placeholderId = source[start + 8];
        }

        /**
         * Gets the comment number (note - each user normally has their own count).
         * @return the comment number.
         */
        public int getPlaceholderId()
        {
            return _placeholderId;
        }

        /**
         * Sets the comment number (note - each user normally has their own count).
         * @param number the comment number.
         */
        public void setPlaceholderId(int number)
        {
            _placeholderId = (byte)number;
        }

        /**
         * Gets the record type.
         * @return the record type.
         */
        public override long GetRecordType() { return RecordTypes.RoundTripHFPlaceholder12.typeID; }

        /**
         * Write the contents of the record back, so it can be written
         * to disk
         *
         * @param out the output stream to write to.
         * @throws java.io.IOException if an error occurs.
         */
        public override void WriteOut(OutputStream _out)
        {
            _out.Write(_header);
            _out.Write(_placeholderId);
        }

        
        public override IDictionary<string, Func<object>> GetGenericProperties()
        {
            throw new NotImplementedException();
        }

        //@Override
        //public Map<String, Supplier<?>> getGenericProperties() {
        //    return GenericRecordUtil.getGenericProperties("placeholderId", this::getPlaceholderId );
        //}
    }
}
