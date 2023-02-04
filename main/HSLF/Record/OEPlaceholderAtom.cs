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
     * OEPlaceholderAtom (3011).<p>
     *
     * An atom record that specifies whether a shape is a placeholder shape.
     *
     * @see Placeholder
     */

    public class OEPlaceholderAtom : RecordAtom
    {

        /**
         * The full size of the master body text placeholder shape.
         */
        public static int PLACEHOLDER_FULLSIZE = 0;

        /**
         * Half of the size of the master body text placeholder shape.
         */
        public static int PLACEHOLDER_HALFSIZE = 1;

        /**
         * A quarter of the size of the master body text placeholder shape.
         */
        public static int PLACEHOLDER_QUARTSIZE = 2;

        private byte[] _header;

        private int placementId;
        private int placeholderId;
        private int placeholderSize;
        private short unusedShort;


        /**
         * Create a new instance of {@code OEPlaceholderAtom}
         */
        public OEPlaceholderAtom()
        {
            _header = new byte[8];
            LittleEndian.PutUShort(_header, 0, 0);
            LittleEndian.PutUShort(_header, 2, (int)GetRecordType());
            LittleEndian.PutInt(_header, 4, 8);

            placementId = 0;
            placeholderId = 0;
            placeholderSize = 0;
        }

        /**
         * Build an instance of {@code OEPlaceholderAtom} from on-disk data
         */
        OEPlaceholderAtom(byte[] source, int start, int len)
        {
            _header = Arrays.CopyOfRange(source, start, start + 8);
            int offset = start + 8;

            placementId = LittleEndian.GetInt(source, offset); offset += 4;
            placeholderId = LittleEndian.GetUByte(source, offset); offset++;
            placeholderSize = LittleEndian.GetUByte(source, offset); offset++;
            unusedShort = LittleEndian.GetShort(source, offset);
        }

        /**
         * @return type of this record {@link RecordTypes#OEPlaceholderAtom}.
         */

        public override long GetRecordType() { return RecordTypes.OEPlaceholderAtom.typeID; }

        /**
         * Returns the placement Id.<p>
         *
         * The placement Id is a number assigned to the placeholder. It goes from -1 to the number of placeholders.
         * It SHOULD be unique among all PlacholderAtom records contained in the corresponding slide.
         * The value 0xFFFFFFFF specifies that the corresponding shape is not a placeholder shape.
         *
         * @return the placement Id.
         */
        public int getPlacementId()
        {
            return placementId;
        }

        /**
         * Sets the placement Id.<p>
         *
         * The placement Id is a number assigned to the placeholder. It goes from -1 to the number of placeholders.
         * It SHOULD be unique among all PlacholderAtom records contained in the corresponding slide.
         * The value 0xFFFFFFFF specifies that the corresponding shape is not a placeholder shape.
         *
         * @param id the placement Id.
         */
        public void setPlacementId(int id)
        {
            placementId = id;
        }

        /**
         * Returns the placeholder Id.<p>
         *
         * placeholder Id specifies the type of the placeholder shape.
         * The value MUST be one of the static constants defined in this class
         *
         * @return the placeholder Id.
         */
        public int getPlaceholderId()
        {
            return placeholderId;
        }

        /**
         * Sets the placeholder Id.<p>
         *
         * placeholder Id specifies the type of the placeholder shape.
         * The value MUST be one of the static constants defined in {@link Placeholder}
         *
         * @param id the placeholder Id.
         */
        public void setPlaceholderId(byte id)
        {
            placeholderId = id;
        }

        /**
         * Returns the placeholder size.
         * Must be one of the PLACEHOLDER_* static constants defined in this class.
         *
         * @return the placeholder size.
         */
        public int getPlaceholderSize()
        {
            return placeholderSize;
        }

        /**
         * Sets the placeholder size.
         * Must be one of the PLACEHOLDER_* static constants defined in this class.
         *
         * @param size the placeholder size.
         */
        public void setPlaceholderSize(byte size)
        {
            placeholderSize = size;
        }

        /**
         * Write the contents of the record back, so it can be written to disk
         */

        public override void WriteOut(OutputStream _out)
        {
            _out.Write(_header);

            byte[] recdata = new byte[8];
            LittleEndian.PutInt(recdata, 0, placementId);
            recdata[4] = (byte)placeholderId;
            recdata[5] = (byte)placeholderSize;
            LittleEndian.PutShort(recdata, 6, unusedShort);

            _out.Write(recdata);
        }

                
        public override IDictionary<string, Func<object>> GetGenericProperties()
        {
            throw new NotImplementedException();
        }


        //public Map<String, Supplier<?>> getGenericProperties() {
        //    return GenericRecordUtil.getGenericProperties(
        //        "placementId", this::getPlacementId,
        //        "placeholderId", this::getPlaceholderId,
        //        "placeholderSize", this::getPlaceholderSize
        //    );
        //}
    }
}