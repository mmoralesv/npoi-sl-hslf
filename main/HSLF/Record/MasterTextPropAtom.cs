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
using NPOI.HSLF.Model;
using NPOI.Util;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace NPOI.HSLF.Record
{
    /**
     * Specifies the Indent Level for the text
     */
    public class MasterTextPropAtom : RecordAtom
    {

        //arbitrarily selected; may need to increase
        private static int DEFAULT_MAX_RECORD_LENGTH = 100_000;
        private static int MAX_RECORD_LENGTH = DEFAULT_MAX_RECORD_LENGTH;

        /**
         * Record header.
         */
        private byte[] _header;

        /**
         * Record data.
         */
        private byte[] _data;

        // indent details
        private List<IndentProp> indents;

        /**
         * @param length the max record length allowed for MasterTextPropAtom
         */
        public static void setMaxRecordLength(int length)
        {
            MAX_RECORD_LENGTH = length;
        }

        /**
         * @return the max record length allowed for MasterTextPropAtom
         */
        public static int getMaxRecordLength()
        {
            return MAX_RECORD_LENGTH;
        }

        /**
         * Constructs a new empty master text prop atom.
         */
        public MasterTextPropAtom()
        {
            _header = new byte[8];
            _data = new byte[0];

            LittleEndian.PutShort(_header, 2, (short)GetRecordType());
            LittleEndian.PutInt(_header, 4, _data.Length);

            indents = new List<IndentProp>();
        }

        /**
         * Constructs the ruler atom record from its
         *  source data.
         *
         * @param source the source data as a byte array.
         * @param start the start offset into the byte array.
         * @param len the length of the slice in the byte array.
         */
        protected MasterTextPropAtom(byte[] source, int start, int len)
        {
            // Get the header.
            _header = Arrays.CopyOfRange(source, start, start + 8);

            // Get the record data.
            _data = IOUtils.SafelyClone(source, start + 8, len - 8, MAX_RECORD_LENGTH);

            try
            {
                read();
            }
            catch (Exception e)
            {
                //LOG.atError().withThrowable(e).log("Failed to parse MasterTextPropAtom");
            }
        }

        /**
         * Gets the record type.
         *
         * @return the record type.
         */

        public override long GetRecordType()
        {
            return RecordTypes.MasterTextPropAtom.typeID;
        }

        /**
         * Write the contents of the record back, so it can be written
         * to disk.
         *
         * @param out the output stream to write to.
         * @throws java.io.IOException if an error occurs.
         */

        public override void WriteOut(OutputStream _out)
        {
            Write();
            _out.Write(_header);
            _out.Write(_data);
        }

        /**
         * Write the internal variables to the record bytes
         */
        private void Write()
        {
            int pos = 0;
            long newSize = Math.BigMul((int)indents.Count, (int)6);
            _data = IOUtils.SafelyAllocate(newSize, MAX_RECORD_LENGTH);
            foreach (IndentProp prop in indents)
            {
                LittleEndian.PutInt(_data, pos, prop.getCharactersCovered());
                LittleEndian.PutShort(_data, pos + 4, (short)prop.getIndentLevel());
                pos += 6;
            }
        }

        /**
         * Read the record bytes and initialize the internal variables
         */
        private void read()
        {
            int pos = 0;
            indents = new List<IndentProp>(_data.Length / 6);

            while (pos <= _data.Length - 6)
            {
                int count = LittleEndian.GetInt(_data, pos);
                short indent = LittleEndian.GetShort(_data, pos + 4);
                indents.Add(new IndentProp(count, indent));
                pos += 6;
            }
        }

        /**
         * Returns the indent that applies at the given text offset
         */
        public int getIndentAt(int offset)
        {
            int charsUntil = 0;
            foreach (IndentProp prop in indents)
            {
                charsUntil += prop.getCharactersCovered();
                if (offset < charsUntil)
                {
                    return prop.getIndentLevel();
                }
            }
            return -1;
        }

        public List<IndentProp> getIndents()
        {
            return indents;
        }


        public IDictionary<string, Func<T>> GetGenericProperties<T>()
        {
            return (IDictionary<string, Func<T>>)GenericRecordUtil.GetGenericProperties(
                "indents", getIndents
            );
        }

        public override IDictionary<string, Func<object>> GetGenericProperties()
        {
            return (IDictionary<string, Func<object>>) GenericRecordUtil.GetGenericProperties(
				"indents", getIndents
			);
		}
    }
}
