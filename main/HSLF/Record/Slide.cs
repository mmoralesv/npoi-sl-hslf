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
namespace NPOI.HSLF.Record
{
    using NPOI.Util;
    using System.Collections.Generic;
    using System.Linq;
    
    /**
     * Master container for Slides. There is one of these for every slide,
     *  and they have certain specific children
     */

    public class Slide : SheetContainer
    {
        private byte[] _header;
        private static long _type = 1006L;

        // Links to our more interesting children
        private SlideAtom slideAtom;
        private PPDrawing ppDrawing;
        private ColorSchemeAtom _colorScheme;

        /**
     * Returns the SlideAtom of this Slide
     */
        public SlideAtom GetSlideAtom()
        {
            return slideAtom;
        }

        /**
     * Returns the PPDrawing of this Slide, which has all the
     *  interesting data in it
     */
        public override PPDrawing GetPPDrawing()
        {
            return ppDrawing;
        }

        /**
     * Set things up, and find our more interesting children
     */
        protected Slide(byte[] source, int start, int len)
        {
            // Grab the header
            _header = Arrays.CopyOfRange(source, start, start+8);

            // Find our children
            _children = FindChildRecords(source, start + 8, len - 8);

            // Find the interesting ones in there
            foreach (Record child in _children)
            {
                if (child is SlideAtom)
                {
                    slideAtom = (SlideAtom)child;
                }
                else if (child is PPDrawing)
                {
                    ppDrawing = (PPDrawing)child;
                }

                if (ppDrawing != null && child is ColorSchemeAtom)
                {
                    _colorScheme = (ColorSchemeAtom)child;
                }
            }
        }

        /**
     * Create a new, empty, Slide, along with its required
     *  child records.
     */
        public Slide()
        {
            _header = new byte[8];
            LittleEndian.PutUShort(_header, 0, 15);
            LittleEndian.PutUShort(_header, 2, (int)_type);
            LittleEndian.PutInt(_header, 4, 0);

            slideAtom = new SlideAtom();
            ppDrawing = new PPDrawing();

            ColorSchemeAtom colorAtom = new ColorSchemeAtom();

            _children = new Record[]
            {
                slideAtom,
                ppDrawing,
                colorAtom
            };
        }

        /**
     * We are of type 1006
     */
        public override long GetRecordType()
        {
            return _type;
        }

        /**
     * Write the contents of the record back, so it can be written
     *  to disk
     */
        public override void WriteOut(OutputStream output)
        {
            WriteOut(_header[0],_header[1],_type,_children,output);
        }

        public override ColorSchemeAtom GetColorScheme()
        {
            return _colorScheme;
        }
    }
}
