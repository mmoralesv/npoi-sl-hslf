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
     * An atomic record containing information about a comment.
     */

    public class Comment2000Atom : RecordAtom
    {

        /**
     * Record header.
     */
        private byte[] _header;

        /**
     * Record data.
     */
        private byte[] _data;

        /**
     * Constructs a brand new comment atom record.
     */
        public Comment2000Atom()
        {
            _header = new byte[8];
            _data = new byte[28];

            LittleEndian.PutShort(_header, 2, (short)GetRecordType());
            LittleEndian.PutInt(_header, 4, _data.Length);

            // It is fine for the other values to be zero
        }

        /**
     * Constructs the comment atom record from its source data.
     *
     * @param source the source data as a byte array.
     * @param start the start offset into the byte array.
     * @param len the length of the slice in the byte array.
     */
        protected Comment2000Atom(byte[] source, int start, int len)
        {
            // Get the header.
            _header = Arrays.CopyOfRange(source, start, start + 8);

            // Get the record data.
            _data = IOUtils.SafelyClone(source, start + 8, len - 8, GetMaxRecordLength());
        }

        /**
     * Gets the comment number (note - each user normally has their own count).
     * @return the comment number.
     */
        public int GetNumber()
        {
            return LittleEndian.GetInt(_data, 0);
        }

        /**
     * Sets the comment number (note - each user normally has their own count).
     * @param number the comment number.
     */
        public void SetNumber(int number)
        {
            LittleEndian.PutInt(_data, 0, number);
        }

        /**
     * Gets the date the comment was made.
     * @return the comment date.
     */
        public DateTime GetDate()
        {
            //return SystemTimeUtils.getDate(_data, 4);
            return new DateTime();
        }

        /**
     * Sets the date the comment was made.
     * @param date the comment date.
     */
        public void SetDate(DateTime date)
        {
            //SystemTimeUtils.storeDate(date, _data, 4);
        }

        /**
     * Gets the X offset of the comment on the page.
     * @return the X offset.
     */
        public int GetXOffset()
        {
            return LittleEndian.GetInt(_data, 20);
        }

        /**
     * Sets the X offset of the comment on the page.
     * @param xOffset the X offset.
     */
        public void SetXOffset(int xOffset)
        {
            LittleEndian.PutInt(_data, 20, xOffset);
        }

        /**
     * Gets the Y offset of the comment on the page.
     * @return the Y offset.
     */
        public int GetYOffset()
        {
            return LittleEndian.GetInt(_data, 24);
        }

        /**
     * Sets the Y offset of the comment on the page.
     * @param yOffset the Y offset.
     */
        public void SetYOffset(int yOffset)
        {
            LittleEndian.PutInt(_data, 24, yOffset);
        }

        /**
     * Gets the record type.
     * @return the record type.
     */
        public override long GetRecordType()
        {
            return RecordTypes.Comment2000Atom.typeID;
        }

        /**
     * Write the contents of the record back, so it can be written
     * to disk
     *
     * @param out the output stream to write to.
     * @throws IOException if an error occurs.
     */
        public override void WriteOut(OutputStream outputStream)
        {
            try
            {
                outputStream.Write(_header);
                outputStream.Write(_data);
            }
            catch (Exception e)
            {
                throw new IOException();
            }
        }

        // public Map<String, Supplier<?>> getGenericProperties() {
        //     return GenericRecordUtil.getGenericProperties(
        //         "number", this::getNumber,
        //         "date", this::getDate,
        //         "xOffset", this::getXOffset,
        //         "yOffset", this::getYOffset
        //     );
        // }

        public override IDictionary<string, Func<object>> GetGenericProperties()
        {
            throw new NotImplementedException();
        }
    }
}
