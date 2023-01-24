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
using NPOI.Util;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http.Headers;

namespace NPOI.HSLF.Record
{

    /**
     * Container records which always exists inside Document.
     * It always acts as a holder for escher DGG container
     *  which may contain which Escher BStore container information
     *  about pictures containes in the presentation (if any).
     */
    public sealed class PPDrawingGroup : RecordAtom
    {

        //arbitrarily selected; may need to increase
        private const int MAX_RECORD_LENGTH = 10_485_760;


        private byte[] _header;
        private EscherContainerRecord dggContainer;
        //cached dgg
        private EscherDggRecord dgg;

        PPDrawingGroup(byte[] source, int start, int len)
        {
            // Get the header
            _header = source.Skip(start).Take(8).ToArray();

            // Get the contents for now
            byte[] contents = IOUtils.SafelyClone(source, start, len, MAX_RECORD_LENGTH);

            DefaultEscherRecordFactory erf = new DefaultEscherRecordFactory();
            EscherRecord child = erf.CreateRecord(contents, 0);
            child.FillFields(contents, 0, erf);
            dggContainer = (EscherContainerRecord)child.GetChild(0);
        }

        /**
         * We are type 1035
         */
        public override long GetRecordType()
        {
            return RecordTypes.PPDrawingGroup.typeID;
        }

        /**
         * We're pretending to be an atom, so return null
         */
        public override Record[] GetChildRecords()
        {
            return null;
        }

        public override void WriteOut(OutputStream os)
        {
            // TODO
        }

        public EscherContainerRecord getDggContainer()
        {
            return dggContainer;
        }

        public EscherDggRecord getEscherDggRecord()
        {
            if (dgg == null)
            {
                foreach (EscherRecord r in dggContainer.GetChildIterator())
                {
                    if (r is EscherDggRecord record)
                    {
                        dgg = record;
                        break;
                    }
                }
            }
            return dgg;
        }

        public override IDictionary<string, Func<object>> GetGenericProperties()
        {
            return new Dictionary<string, Func<object>> {
            { "dggContainer", getDggContainer }
        };
        }
    }
}