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
     * A SlideShowSlideInfo Atom (type 1017).<br>
     * <br>
     *
     * An atom record that specifies which transition effects to perform
     * during a slide show, and how to advance to the next presentation slide.<br>
     * <br>
     *
     * Combination of effectType and effectDirection:
     * <table>
     * <caption>Combination of effectType and effectDirection</caption>
     * <tr><th>type</th><th>description</th><th>direction</th></tr>
     * <tr><td>0</td><td>cut</td><td>0x00 = no transition, 0x01 = black transition</td></tr>
     * <tr><td>1</td><td>random</td><td>0x00</td></tr>
     * <tr><td>2</td><td>blinds</td><td>0x00 = vertical, 0x01 = horizontal</td></tr>
     * <tr><td>3</td><td>checker</td><td>like blinds</td></tr>
     * <tr><td>4</td><td>cover</td><td>0x00 = left, 0x01 = up, 0x02 = right, 0x03 = down, 0x04 = left/up, 0x05 = right/up, 0x06 left/down, 0x07 = left/down</td></tr>
     * <tr><td>5</td><td>dissolve</td><td>0x00</td></tr>
     * <tr><td>6</td><td>fade</td><td>0x00</td></tr>
     * <tr><td>7</td><td>uncover</td><td>like cover</td></tr>
     * <tr><td>8</td><td>random bars</td><td>like blinds</td></tr>
     * <tr><td>9</td><td>strips</td><td>like 0x04 - 0x07 of cover</td></tr>
     * <tr><td>10</td><td>wipe</td><td>like 0x00 - 0x03 of cover</td></tr>
     * <tr><td>11</td><td>box in/out</td><td>0x00 = out, 0x01 = in</td></tr>
     * <tr><td>13</td><td>split</td><td>0x00 = horizontally out, 0x01 = horizontally in, 0x02 = vertically out, 0x03 = vertically in</td></tr>
     * <tr><td>17</td><td>diamond</td><td>0x00</td></tr>
     * <tr><td>18</td><td>plus</td><td>0x00</td></tr>
     * <tr><td>19</td><td>wedge</td><td>0x00</td></tr>
     * <tr><td>20</td><td>push</td><td>like 0x00 - 0x03 of cover</td></tr>
     * <tr><td>21</td><td>comb</td><td>like blinds</td></tr>
     * <tr><td>22</td><td>newsflash</td><td>0x00</td></tr>
     * <tr><td>23</td><td>alphafade</td><td>0x00</td></tr>
     * <tr><td>26</td><td>wheel</td><td>number of radial divisions (0x01,0x02,0x03,0x04,0x08)</td></tr>
     * <tr><td>27</td><td>circle</td><td>0x00</td></tr>
     * <tr><td>255</td><td>undefined</td><td>0x00</td></tr>
     * </table>
     */
    public class SSSlideInfoAtom : RecordAtom
    {
        /**
         * A bit that specifies whether the presentation slide can be
         * manually advanced by the user during the slide show.
         */
        public static int MANUAL_ADVANCE_BIT = 1 << 0;

        /**
         * A bit that specifies whether the corresponding slide is
         * hidden and is not displayed during the slide show.
         */
        public static int HIDDEN_BIT = 1 << 2;

        /**
         * A bit that specifies whether to play the sound specified by soundIfRef.
         */
        public static int SOUND_BIT = 1 << 4;

        /**
         * A bit that specifies whether the sound specified by soundIdRef is
         * looped continuously when playing until the next sound plays.
         */
        public static int LOOP_SOUND_BIT = 1 << 6;

        /**
         * A bit that specifies whether to stop any currently playing
         * sound when the transition starts.
         */
        public static int STOP_SOUND_BIT = 1 << 8;

        /**
         * A bit that specifies whether the slide will automatically
         * advance after slideTime milliseconds during the slide show.
         */
        public static int AUTO_ADVANCE_BIT = 1 << 10;

        /**
         * A bit that specifies whether to display the cursor during
         * the slide show.
         */
        public static int CURSOR_VISIBLE_BIT = 1 << 12;

        // public static int RESERVED1_BIT       = 1 << 1;
        // public static int RESERVED2_BIT       = 1 << 3;
        // public static int RESERVED3_BIT       = 1 << 5;
        // public static int RESERVED4_BIT       = 1 << 7;
        // public static int RESERVED5_BIT       = 1 << 9;
        // public static int RESERVED6_BIT       = 1 << 11;
        // public static int RESERVED7_BIT       = 1 << 13 | 1 << 14 | 1 << 15;

        private static int[] EFFECT_MASKS = {
        MANUAL_ADVANCE_BIT,
        HIDDEN_BIT,
        SOUND_BIT,
        LOOP_SOUND_BIT,
        STOP_SOUND_BIT,
        AUTO_ADVANCE_BIT,
        CURSOR_VISIBLE_BIT
    };

        private static String[] EFFECT_NAMES = {
        "MANUAL_ADVANCE",
        "HIDDEN",
        "SOUND",
        "LOOP_SOUND",
        "STOP_SOUND",
        "AUTO_ADVANCE",
        "CURSOR_VISIBLE"
    };

        private static long _type = RecordTypes.SSSlideInfoAtom.typeID;

        private byte[] _header;

        /**
         * A signed integer that specifies an amount of time, in milliseconds, to wait
         * before advancing to the next presentation slide. It MUST be greater than or equal to 0 and
         * less than or equal to 86399000. It MUST be ignored unless AUTO_ADVANCE_BIT is TRUE.
         */
        private int _slideTime;

        /**
         * A SoundIdRef that specifies which sound to play when the transition starts.
         */
        private int _soundIdRef;

        /**
         * A byte that specifies the variant of effectType. In combination of the effectType
         * there are further restriction and specification of this field.
         */
        private short _effectDirection; // byte

        /**
         * A byte that specifies which transition is used when transitioning to the
         * next presentation slide during a slide show. Exact rendering of any transition is
         * determined by the rendering application. As such, the same transition can have
         * many variations depending on the implementation.
         */
        private short _effectType; // byte

        /**
         * Various flags - see bitmask for more details
         */
        private int _effectTransitionFlags;

        /**
         * A byte value that specifies how long the transition takes to run.
         * (0x00 = 0.75 seconds, 0x01 = 0.5 seconds, 0x02 = 0.25 seconds)
         */
        private short _speed; // byte
        private byte[] _unused; // 3-byte

        public SSSlideInfoAtom()
        {
            _header = new byte[8];
            LittleEndian.PutShort(_header, 0, (short)0);
            LittleEndian.PutShort(_header, 2, (short)_type);
            LittleEndian.PutShort(_header, 4, (short)0x10);
            LittleEndian.PutShort(_header, 6, (short)0);
            _unused = new byte[3];
        }

        public SSSlideInfoAtom(byte[] source, int offset, int len)
        {
            int ofs = offset;

            // Sanity Checking
            if (len != 24) len = 24;

            if (source.Length < offset + len)
            {
                throw new ArgumentException("Need at least " + (offset + len) +
                        " bytes with offset " + offset + ", length " + len + " and array-size " + source.Length);
            }

            // Get the header
            _header = Arrays.CopyOfRange(source, ofs, ofs + 8);
            ofs += _header.Length;

            if (LittleEndian.GetShort(_header, 0) != 0)
            {
                //LOG.atDebug().log("Invalid data for SSSlideInfoAtom at offset 0: " + LittleEndian.getShort(_header, 0));
            }
            if (LittleEndian.GetShort(_header, 2) != RecordTypes.SSSlideInfoAtom.typeID)
            {
                //LOG.atDebug().log("Invalid data for SSSlideInfoAtom at offset 2: "+ LittleEndian.getShort(_header, 2));
            }
            if (LittleEndian.GetShort(_header, 4) != 0x10)
            {
                //LOG.atDebug().log("Invalid data for SSSlideInfoAtom at offset 4: "+ LittleEndian.getShort(_header, 4));
            }
            if (LittleEndian.GetShort(_header, 6) == 0)
            {
                //LOG.atDebug().log("Invalid data for SSSlideInfoAtom at offset 6: "+ LittleEndian.getShort(_header, 6));
            }

            _slideTime = LittleEndian.GetInt(source, ofs);
            if (_slideTime < 0 || _slideTime > 86399000)
            {
                //LOG.atDebug().log("Invalid data for SSSlideInfoAtom - invalid slideTime: "+ _slideTime);
            }
            ofs += LittleEndianConsts.INT_SIZE;
            _soundIdRef = LittleEndian.GetInt(source, ofs);
            ofs += LittleEndianConsts.INT_SIZE;
            _effectDirection = LittleEndian.GetUByte(source, ofs);
            ofs += LittleEndianConsts.BYTE_SIZE;
            _effectType = LittleEndian.GetUByte(source, ofs);
            ofs += LittleEndianConsts.BYTE_SIZE;
            _effectTransitionFlags = LittleEndian.GetShort(source, ofs);
            ofs += LittleEndianConsts.SHORT_SIZE;
            _speed = LittleEndian.GetUByte(source, ofs);
            ofs += LittleEndianConsts.BYTE_SIZE;
            _unused = Arrays.CopyOfRange(source, ofs, ofs + 3);
        }

        /**
         * Write the contents of the record back, so it can be written
         *  to disk
         */

        public override void WriteOut(OutputStream _out)
        {
            // Header - size or type unchanged
            _out.Write(_header);
            WriteLittleEndian(_slideTime, _out);
            WriteLittleEndian(_soundIdRef, _out);

            byte[] byteBuf = new byte[LittleEndianConsts.BYTE_SIZE];
            LittleEndian.PutUByte(byteBuf, 0, _effectDirection);
            _out.Write(byteBuf);
            LittleEndian.PutUByte(byteBuf, 0, _effectType);
            _out.Write(byteBuf);

            WriteLittleEndian(_effectTransitionFlags, _out);
            LittleEndian.PutUByte(byteBuf, 0, _speed);
            _out.Write(byteBuf);

            if (_unused.Length == 3) _out.Write(_unused);
        }

        /**
         * We are of type 1017
         */

        public override long GetRecordType() { return _type; }


        public int GetSlideTime()
        {
            return _slideTime;
        }

        public void SetSlideTime(int slideTime)
        {
            this._slideTime = slideTime;
        }

        public int GetSoundIdRef()
        {
            return _soundIdRef;
        }

        public void setSoundIdRef(int soundIdRef)
        {
            this._soundIdRef = soundIdRef;
        }

        public short getEffectDirection()
        {
            return _effectDirection;
        }

        public void setEffectDirection(short effectDirection)
        {
            this._effectDirection = effectDirection;
        }

        public short getEffectType()
        {
            return _effectType;
        }

        public void setEffectType(short effectType)
        {
            this._effectType = effectType;
        }

        public int getEffectTransitionFlags()
        {
            return _effectTransitionFlags;
        }

        public void setEffectTransitionFlags(short effectTransitionFlags)
        {
            this._effectTransitionFlags = effectTransitionFlags;
        }

        /**
         * Use one of the bitmasks MANUAL_ADVANCE_BIT ... CURSOR_VISIBLE_BIT
         */
        public void setEffectTransitionFlagByBit(int bitmask, bool enabled)
        {
            if (enabled)
            {
                _effectTransitionFlags |= bitmask;
            }
            else
            {
                _effectTransitionFlags &= (0xFFFF ^ bitmask);
            }
        }

        public bool getEffectTransitionFlagByBit(int bitmask)
        {
            return ((_effectTransitionFlags & bitmask) != 0);
        }

        public short getSpeed()
        {
            return _speed;
        }

        public void setSpeed(short speed)
        {
            this._speed = speed;
        }

        public IDictionary<string, Func<T>> GetGenericProperties<T>()
        {
            return (IDictionary<string, Func<T>>)GenericRecordUtil.GetGenericProperties(
                "effectTransitionFlags", GenericRecordUtil.GetBitsAsString(getEffectTransitionFlags, EFFECT_MASKS, EFFECT_NAMES),
                "slideTime", () => GetSlideTime(),
                "soundIdRef", () => GetSoundIdRef(),
                "effectDirection", () => getEffectDirection(),
                "effectType", () => getEffectType(),
                "speed", () => getSpeed()
            );
        }

        public override IDictionary<string, Func<object>> GetGenericProperties()
        {
            throw new NotImplementedException();
        }
    }
}
