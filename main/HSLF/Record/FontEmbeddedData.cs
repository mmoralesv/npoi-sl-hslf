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

using NPOI.Common.UserModel.Fonts;
using NPOI.Util;
using System;
using System.Collections.Generic;

namespace NPOI.HSLF.Record
{
	public class FontEmbeddedData : RecordAtom, IFontFacet
	{
		private static int DEFAULT_MAX_RECORD_LENGTH = 5_000_000;
		private static int MAX_RECORD_LENGTH = DEFAULT_MAX_RECORD_LENGTH;

		/**
		 * Record header.
		 */
		private byte[] _header;

		/**
		 * Record data - An EOT Font
		 */
		private byte[] _data;

		/**
		 * A cached FontHeader so that we don't keep creating new FontHeader instances
		 */
		private FontHeader fontHeader;

		/**
		 * @return the max record length allowed for FontEmbeddedData
		 */
		public static new int GetMaxRecordLength()
		{
			return MAX_RECORD_LENGTH;
		}

		/**
		 * Constructs a brand new font embedded record.
		 */
		/* package */
		FontEmbeddedData()
		{
			_header = new byte[8];
			_data = new byte[4];

			LittleEndian.PutShort(_header, 2, (short)GetRecordType());
			LittleEndian.PutInt(_header, 4, _data.Length);
		}

		/**
		 * Constructs the font embedded record from its source data.
		 *
		 * @param source the source data as a byte array.
		 * @param start the start offset into the byte array.
		 * @param len the length of the slice in the byte array.
		 */
		/* package */
		FontEmbeddedData(byte[] source, int start, int len)
		{
			// Get the header.
			_header = Arrays.CopyOfRange(source, start, start + 8);

			// Get the record data.
			_data = IOUtils.SafelyClone(source, start + 8, len - 8, MAX_RECORD_LENGTH);

			// Must be at least 4 bytes long
			if (_data.Length < 4)
			{
				throw new InvalidOperationException("The length of the data for a ExObjListAtom must be at least 4 bytes, but was only " + _data.Length);
			}
		}


		public override long GetRecordType()
		{
			return RecordTypes.FontEmbeddedData.typeID;
		}


		public override void WriteOut(OutputStream _out)
		{
			_out.Write(_header);
			_out.Write(_data);
		}

		/**
		 * Overwrite the font data. Reading values from this FontEmbeddedData instance while calling setFontData
		 * is not thread safe.
		 * @param fontData new font data
		 */
		public void SetFontData(byte[] fontData)
		{
			fontHeader = null;
			_data = (byte[])fontData.Clone();
			LittleEndian.PutInt(_header, 4, _data.Length);
		}

		/**
		 * Read the font data. Reading values from this FontEmbeddedData instance while calling {@link #setFontData(byte[])}
		 * is not thread safe.
		 * @return font data
		 */
		public FontHeader GetFontHeader()
		{
			if (fontHeader == null)
			{
				FontHeader h = new FontHeader();
				h.init(_data, 0, _data.Length);
				fontHeader = h;
			}
			return fontHeader;
		}


		public int GetWeight()
		{
			return GetFontHeader().GetWeight();
		}


		public bool IsItalic()
		{
			return GetFontHeader().isItalic();
		}

		public String GetTypeface()
		{
			return GetFontHeader().getFamilyName();
		}


		public Object GetFontData()
		{
			return this;
		}


		public override IDictionary<string, Func<object>> GetGenericProperties()
		{
			return (IDictionary<string, Func<object>>)GenericRecordUtil.GetGenericProperties("fontHeader", GetFontHeader);
		}
	}
}
