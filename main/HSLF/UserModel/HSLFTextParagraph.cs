﻿using NPOI.SL.UserModel;
using NPOI.HSLF.Record;
using NPOI.HSLF.Model;
using System;
using System.Collections.Generic;
using System.Text;
using static NPOI.HSLF.Model.TextPropCollection;
using NPOI.Util;
using NPOI.Common.UserModel.Fonts;
using System.Linq;
using NPOI.HSLF.Exceptions;
using System.Drawing;
using System.IO;

namespace NPOI.HSLF.UserModel
{
	public class HSLFTextParagraph : TextParagraph<HSLFShape, HSLFTextParagraph, HSLFTextRun>
	{
		private static int MAX_NUMBER_OF_STYLES = 20_000;

		// Note: These fields are protected to help with unit testing
		// Other classes shouldn't really go playing with them!
		private TextHeaderAtom _headerAtom;
		private TextBytesAtom _byteAtom;
		private TextCharsAtom _charAtom;
		private TextPropCollection _paragraphStyle;

		protected TextRulerAtom _ruler;
		protected List<HSLFTextRun> _runs = new List<HSLFTextRun>();
		protected HSLFTextShape _parentShape;
		private HSLFSheet _sheet;
		private int shapeId;

		private StyleTextProp9Atom styleTextProp9Atom;

		private bool _dirty;

		private List<HSLFTextParagraph> parentList;

		//public TextShape<HSLFShape, HSLFTextParagraph> ParentShape => throw new NotImplementedException();

		//private class HSLFTabStopDecorator : TabStop
		//{
		//	HSLFTabStop tabStop;

		//	HSLFTabStopDecorator(HSLFTabStop tabStop)
		//	{
		//		this.tabStop = tabStop;
		//	}

		//	//@Override
		//	public double GetPositionInPoints()
		//	{
		//		return tabStop.GetPositionInPoints();
		//	}

		//	//@Override
		//	public void SetPositionInPoints(double position)
		//	{
		//		tabStop.SetPositionInPoints(position);
		//		SetDirty();
		//	}

		//	//@Override
		//	public new TabStopType GetType()
		//	{
		//		return tabStop.GetType();
		//	}

		//	//@Override
		//	public void SetType(TabStopType type)
		//	{
		//		tabStop.SetType(type);
		//		SetDirty();
		//	}


		//}


		/**
		* Constructs a Text Run from a Unicode text block.
		* Either a {@link TextCharsAtom} or a {@link TextBytesAtom} needs to be provided.
		 *
		* @param tha the TextHeaderAtom that defines what's what
		* @param tba the TextBytesAtom containing the text or null if {@link TextCharsAtom} is provided
		* @param tca the TextCharsAtom containing the text or null if {@link TextBytesAtom} is provided
		* @param parentList the list which contains this paragraph
		 */
		/* package */
		public HSLFTextParagraph(TextHeaderAtom tha, TextBytesAtom tba, TextCharsAtom tca, List<HSLFTextParagraph> parentList)
		{
			if (tha == null)
			{
				throw new InvalidOperationException("TextHeaderAtom must be set.");
			}
			_headerAtom = tha;
			_byteAtom = tba;
			_charAtom = tca;
			this.parentList = parentList;
			SetParagraphStyle(new TextPropCollection(1, TextPropType.paragraph));
		}

		public void AddTextRun(HSLFTextRun run)
		{
			_runs.Add(run);
		}

		//@Override
		public List<HSLFTextRun> GetTextRuns()
		{
			return _runs;
		}

		public TextPropCollection GetParagraphStyle()
		{
			return _paragraphStyle;
		}

		public void SetParagraphStyle(TextPropCollection paragraphStyle)
		{
			_paragraphStyle = paragraphStyle;
		}

		/**
		 * Supply the Sheet we belong to, which might have an assigned SlideShow
		 * Also passes it on to our child RichTextRuns
		 */
		//public static void SupplySheet(List<HSLFTextParagraph> paragraphs, HSLFSheet sheet)
		//{
		//	if (paragraphs == null)
		//	{
		//		return;
		//	}
		//	foreach (HSLFTextParagraph p in paragraphs)
		//	{
		//		p.SupplySheet(sheet);
		//	}

		//	//assert(sheet.getSlideShow() != null);
		//}

		/**
		 * Supply the Sheet we belong to, which might have an assigned SlideShow
		 * Also passes it on to our child RichTextRuns
		 */
		//private void SupplySheet(HSLFSheet sheet)
		//{
		//	this._sheet = sheet;

		//	foreach (HSLFTextRun rt in _runs)
		//	{
		//		rt.UpdateSheet();
		//	}
		//}

		public HSLFSheet GetSheet()
		{
			return this._sheet;
		}

		/**
		 * @return Shape ID
		 */
		protected int GetShapeId()
		{
			return shapeId;
		}

		/**
		 * @param id Shape ID
		 */
		protected void SetShapeId(int id)
		{
			shapeId = id;
		}

		/**
		 * @return 0-based index of the text run in the SLWT container
		 */
		public int GetIndex()
		{
			return (_headerAtom != null) ? _headerAtom.GetIndex() : -1;
		}

		/**
		 * Sets the index of the paragraph in the SLWT container
		 */
		protected void SetIndex(int index)
		{
			if (_headerAtom != null)
			{
				_headerAtom.SetIndex(index);
			}
		}

		/**
		 * Returns the type of the text, from the TextHeaderAtom.
		 * Possible values can be seen from TextHeaderAtom
		 * @see TextHeaderAtom
		 */
		public int GetRunType()
		{
			return (_headerAtom != null) ? _headerAtom.GetTextType() : -1;
		}

		public void SetRunType(int runType)
		{
			if (_headerAtom != null)
			{
				_headerAtom.SetTextType(runType);
			}
		}

		/**
		 * Is this Text Run one from a {@link PPDrawing}, or is it
		 *  one from the {@link SlideListWithText}?
		 */
		public bool IsDrawingBased()
		{
			return (GetIndex() == -1);
		}

		public TextRulerAtom GetTextRuler()
		{
			return _ruler;
		}

		public TextRulerAtom CreateTextRuler()
		{
			_ruler = GetTextRuler();
			if (_ruler == null)
			{
				_ruler = TextRulerAtom.GetParagraphInstance();
				Record.Record childAfter = _byteAtom;
				if (childAfter == null)
				{
					childAfter = _charAtom;
				}
				if (childAfter == null)
				{
					childAfter = _headerAtom;
				}
				_headerAtom.GetParentRecord().AddChildAfter(_ruler, childAfter);
			}
			return _ruler;
		}

		/**
		 * Returns records that make up the list of text paragraphs
		 * (there can be misc InteractiveInfo, TxInteractiveInfo and other records)
		 *
		 * @return text run records
		 */
		public Record.Record[] GetRecords()
		{
			Record.Record[] r = _headerAtom.GetParentRecord().GetChildRecords();
			return GetRecords(r, new int[] { 0 }, _headerAtom);
		}

		private static Record.Record[] GetRecords(Record.Record[] records, int[] startIdx, TextHeaderAtom headerAtom)
		{
			if (records == null)
			{
				throw new NullReferenceException("records need to be set.");
			}

			for (; startIdx[0] < records.Length; startIdx[0]++)
			{
				Record.Record r = records[startIdx[0]];
				if (r is TextHeaderAtom && (headerAtom == null || r == headerAtom))
				{
					break;
				}
			}

			if (startIdx[0] >= records.Length)
			{
				//LOG.atInfo().log("header atom wasn't found - container might contain only an OutlineTextRefAtom");
				return new Record.Record[0];
			}

			int length;
			for (length = 1; startIdx[0] + length < records.Length; length++)
			{
				Record.Record r = records[startIdx[0] + length];
				if (r is TextHeaderAtom || r is SlidePersistAtom)
				{
					break;
				}
			}

			Record.Record[] result = Arrays.CopyOfRange<Record.Record>(records, startIdx[0], startIdx[0] + length);
			startIdx[0] += length;

			return result;
		}

		/** Numbered List info */
		public void SetStyleTextProp9Atom(StyleTextProp9Atom styleTextProp9Atom)
		{
			this.styleTextProp9Atom = styleTextProp9Atom;
		}

		/** Numbered List info */
		public StyleTextProp9Atom GetStyleTextProp9Atom()
		{
			return this.styleTextProp9Atom;
		}

		//@Override
		public IEnumerator<HSLFTextRun> Iterator()
		{
			return _runs.GetEnumerator();
		}

		/**
		 * @since POI 5.2.0
		 */
		//@Override
		public IEnumerator<HSLFTextRun> Spliterator()
		{
			return _runs.GetEnumerator();
		}

		//@Override
		public Double GetLeftMargin()
		{
			int val = 0;
			if (_ruler != null)
			{
				int[] toList = _ruler.GetTextOffsets();
				val = (toList.Length > GetIndentLevel()) ? toList[GetIndentLevel()] : 0;
			}

			if (val == 0)
			{
				TextProp tp = GetPropVal<TextProp>(_paragraphStyle, "text.offset");
				val = (tp == null) ? 0 : tp.GetValue();
			}

			return (val == 0) ? 0 : Units.MasterToPoints(val);
		}

		//@Override
		public void SetLeftMargin(Double leftMargin)
		{
			int val = (leftMargin == 0) ? 0 : Units.PointsToMaster(leftMargin);
			SetParagraphTextPropVal("text.offset", val);
		}

		//@Override
		public Double GetRightMargin()
		{
			// TODO: find out, how to determine this value
			return 0;
		}

		//@Override
		public void SetRightMargin(Double rightMargin)
		{
			// TODO: find out, how to set this value
		}

		//@Override
		public Double GetIndent()
		{
			int val = 0;
			if (_ruler != null)
			{
				int[] toList = _ruler.GetBulletOffsets();
				val = (toList.Length > GetIndentLevel()) ? toList[GetIndentLevel()] : 0;
			}

			if (val == 0)
			{
				TextProp tp = GetPropVal<TextProp>(_paragraphStyle, "bullet.offset");
				val = (tp == null) ? 0 : tp.GetValue();
			}

			return (val == 0) ? 0 : Units.MasterToPoints(val);
		}

		//@Override
		public void SetIndent(Double indent)
		{
			int val = (indent == 0) ? 0 : Units.PointsToMaster(indent);
			SetParagraphTextPropVal("bullet.offset", val);
		}

		//@Override
		public String GetDefaultFontFamily()
		{
			FontInfo fontInfo = null;
			if (!(_runs.Count == 0))
			{
				HSLFTextRun tr = _runs.ElementAt(0);
				fontInfo = tr.GetFontInfo(FontGroupEnum.LATIN);
				// fallback to LATIN if the font for the font group wasn't defined
				if (fontInfo == null)
				{
					fontInfo = (FontInfo)tr.GetFontInfo(FontGroupEnum.LATIN);
				}
			}
			if (fontInfo == null)
			{
				fontInfo = HSLFFontInfoPredefined.ARIAL;
			}

			return fontInfo.GetTypeface();
		}

		//@Override
		public Double GetDefaultFontSize()
		{
			Double d = 0;
			if (!(_runs.Count == 0))
			{
				d = _runs.ElementAt(0).GetFontSize();
			}

			return (d != 0) ? d : 12d;
		}

		//@Override
		public void SetTextAlign(TextAlign align)
		{
			int alignInt = 0;
			if (align != null)
			{
				switch (align)
				{
					default:
					case TextAlign.LEFT: alignInt = TextAlignmentProp.LEFT; break;
					case TextAlign.CENTER: alignInt = TextAlignmentProp.CENTER; break;
					case TextAlign.RIGHT: alignInt = TextAlignmentProp.RIGHT; break;
					case TextAlign.DIST: alignInt = TextAlignmentProp.DISTRIBUTED; break;
					case TextAlign.JUSTIFY: alignInt = TextAlignmentProp.JUSTIFY; break;
					case TextAlign.JUSTIFY_LOW: alignInt = TextAlignmentProp.JUSTIFYLOW; break;
					case TextAlign.THAI_DIST: alignInt = TextAlignmentProp.THAIDISTRIBUTED; break;
				}
			}
			SetParagraphTextPropVal("alignment", alignInt);
		}

		//@Override
		public TextAlign GetTextAlign()
		{
			TextProp tp = GetPropVal<TextProp>(_paragraphStyle, "alignment");
			if (tp == null)
			{
				return TextAlign.LEFT;
			}
			switch (tp.GetValue())
			{
				default:
				case TextAlignmentProp.LEFT: return TextAlign.LEFT;
				case TextAlignmentProp.CENTER: return TextAlign.CENTER;
				case TextAlignmentProp.RIGHT: return TextAlign.RIGHT;
				case TextAlignmentProp.JUSTIFY: return TextAlign.JUSTIFY;
				case TextAlignmentProp.JUSTIFYLOW: return TextAlign.JUSTIFY_LOW;
				case TextAlignmentProp.DISTRIBUTED: return TextAlign.DIST;
				case TextAlignmentProp.THAIDISTRIBUTED: return TextAlign.THAI_DIST;
			}
		}

		//@Override
		public FontAlign GetFontAlign()
		{
			TextProp tp = GetPropVal<TextProp>(_paragraphStyle, FontAlignmentProp.NAME);
			if (tp == null)
			{
				return FontAlign.AUTO;
			}

			switch (tp.GetValue())
			{
				case FontAlignmentProp.BASELINE: return FontAlign.BASELINE;
				case FontAlignmentProp.TOP: return FontAlign.TOP;
				case FontAlignmentProp.CENTER: return FontAlign.CENTER;
				case FontAlignmentProp.BOTTOM: return FontAlign.BOTTOM;
				default: return FontAlign.AUTO;
			}
		}

		public AutoNumberingScheme GetAutoNumberingScheme()
		{
			if (styleTextProp9Atom == null)
			{
				return null;
			}
			TextPFException9[] ant = styleTextProp9Atom.GetAutoNumberTypes();
			int level = GetIndentLevel();
			if (ant == null || level == -1 || level >= ant.Length)
			{
				return null;
			}
			return ant[level].GetAutoNumberScheme();
		}

		public int GetAutoNumberingStartAt()
		{
			if (styleTextProp9Atom == null)
			{
				return 0;
			}
			TextPFException9[] ant = styleTextProp9Atom.GetAutoNumberTypes();
			int level = GetIndentLevel();
			if (ant == null || level >= ant.Length)
			{
				return 0;
			}
			short startAt = ant[level].GetAutoNumberStartNumber();
			//assert(startAt != null);
			return startAt;
		}


		//@Override
		public BulletStyle GetBulletStyle()
		{
			if (!IsBullet() && GetAutoNumberingScheme() == null)
			{
				return null;
			}

			return new BaseBulletStyle(this);
		}

		//@Override
		public void SetBulletStyle(params object[] styles)
		{
			if (styles.Length == 0)
			{
				SetBullet(false);
			}
			else
			{
				SetBullet(true);
				foreach (Object ostyle in styles)
				{
					if (ostyle is Number)
					{
						SetBulletSize(((Number)ostyle).GetDoubleValue());
					}
					else if (ostyle is Color)
					{
						//SetBulletColor((Color)ostyle);
					}
					else if (ostyle is Character)
					{
						SetBulletChar((Character)ostyle);
					}
					else if (ostyle is String)
					{
						//SetBulletFont((String)ostyle);
					}
					else if (ostyle is AutoNumberingScheme)
					{
						throw new HSLFException("setting bullet auto-numberin scheme for HSLF not supported ... yet");
					}
				}
			}
		}

		//@Override
		public HSLFTextShape GetParentShape()
		{
			return _parentShape;
		}

		public void SetParentShape(HSLFTextShape parentShape)
		{
			_parentShape = parentShape;
		}

		//@Override
		public int GetIndentLevel()
		{
			return _paragraphStyle == null ? 0 : _paragraphStyle.GetIndentLevel();
		}

		//@Override
		public void SetIndentLevel(int level)
		{
			if (_paragraphStyle != null)
			{
				_paragraphStyle.setIndentLevel((short)level);
			}
		}

		/**
		 * Sets whether this rich text run has bullets
		 */
		public void SetBullet(bool flag)
		{
			SetFlag(ParagraphFlagsTextProp.BULLET_IDX, flag);
		}

		/**
		 * Returns whether this rich text run has bullets
		 */
		public bool IsBullet()
		{
			return GetFlag(ParagraphFlagsTextProp.BULLET_IDX);
		}

		/**
		 * Sets the bullet character
		 */
		public void SetBulletChar(Character c)
		{
			int val = (c == null) ? 0 : Convert.ToInt32(c);
			SetParagraphTextPropVal("bullet.char", val);
		}

		/**
		 * Returns the bullet character
		 */
		public char GetBulletChar()
		{
			TextProp tp = GetPropVal<TextProp>(_paragraphStyle, "bullet.char");
			return (tp == null) ? ' ' : (char)tp.GetValue();
		}

		/**
		 * Sets the bullet size
		 */
		public void SetBulletSize(Double size)
		{
			SetPctOrPoints("bullet.size", size);
		}

		/**
		 * Returns the bullet size, null if unset
		 */
		public Double GetBulletSize()
		{
			return GetPctOrPoints("bullet.size");
		}

		/**
		 * Sets the bullet color
		 */
		//public void SetBulletColor(Color color)
		//{
		//	int val = (color == null) ? null : new Color(color.getBlue(), color.getGreen(), color.getRed(), 254).getRGB();
		//	setParagraphTextPropVal("bullet.color", val);
		//	setFlag(ParagraphFlagsTextProp.BULLET_HARDCOLOR_IDX, (color != null));
		//}

		/**
		 * Returns the bullet color
		 */
		//public Color GetBulletColor()
		//{
		//	TextProp tp = GetPropVal(_paragraphStyle, "bullet.color");
		//	bool hasColor = GetFlag(ParagraphFlagsTextProp.BULLET_HARDCOLOR_IDX);
		//	if (tp == null || !hasColor)
		//	{
		//		// if bullet color is undefined, return color of first run
		//		if (_runs.Count == 0)
		//		{
		//			return Color.White;
		//		}

		//		SolidPaint sp = _runs.ElementAt(0).GetFontColor();
		//		if (sp == null)
		//		{
		//			return Color.White;
		//		}

		//		return DrawPaint.ApplyColorTransform(sp.getSolidColor());
		//	}

		//	return getColorFromColorIndexStruct(tp.GetValue(), _sheet);
		//}

		/**
		 * Sets the bullet font
		 */
		//public void SetBulletFont(String typeface)
		//{
		//	if (typeface == null)
		//	{
		//		SetPropVal(_paragraphStyle, "bullet.font", 0);
		//		SetFlag(ParagraphFlagsTextProp.BULLET_HARDFONT_IDX, false);
		//		return;
		//	}

		//	HSLFFontInfo fi = new HSLFFontInfo(typeface);
		//	fi = GetSheet().GetSlideShow().AddFont(fi);

		//	SetParagraphTextPropVal("bullet.font", fi.GetIndex());
		//	SetFlag(ParagraphFlagsTextProp.BULLET_HARDFONT_IDX, true);
		//}

		/**
		 * Returns the bullet font
		 */
		public String GetBulletFont()
		{
			TextProp tp = GetPropVal<TextProp>(_paragraphStyle, "bullet.font");
			bool hasFont = GetFlag(ParagraphFlagsTextProp.BULLET_HARDFONT_IDX);
			if (tp == null || !hasFont)
			{
				return GetDefaultFontFamily();
			}
			HSLFFontInfo ppFont = GetSheet().GetSlideShow().GetFont(tp.GetValue());
			//assert(ppFont != null);
			return ppFont.GetTypeface();
		}

		//@Override
		public void SetLineSpacing(Double lineSpacing)
		{
			SetPctOrPoints("linespacing", lineSpacing);
		}

		//@Override
		public Double GetLineSpacing()
		{
			return GetPctOrPoints("linespacing");
		}

		//@Override
		public void SetSpaceBefore(Double spaceBefore)
		{
			SetPctOrPoints("spacebefore", spaceBefore);
		}

		//@Override
		public Double GetSpaceBefore()
		{
			return GetPctOrPoints("spacebefore");
		}

		//@Override
		public void SetSpaceAfter(Double spaceAfter)
		{
			SetPctOrPoints("spaceafter", spaceAfter);
		}

		//@Override
		public Double GetSpaceAfter()
		{
			return GetPctOrPoints("spaceafter");
		}

		//@Override
		public Double GetDefaultTabSize()
		{
			// TODO: implement
			return 0;
		}


		//@Override
		public List<T> GetTabStops<T>() where T : TabStop
		{
			List<HSLFTabStop> tabStops;
			if (GetSheet() is HSLFSlideMaster)
			{
				HSLFTabStopPropCollection tpc = GetMasterPropVal<HSLFTabStopPropCollection>(_paragraphStyle, HSLFTabStopPropCollection.NAME);
				if (tpc == null)
				{
					return null;
				}
				tabStops = tpc.GetTabStops();
			}
			else
			{
				TextRulerAtom textRuler = (TextRulerAtom)_headerAtom.GetParentRecord().FindFirstOfType(RecordTypes.TextRulerAtom.typeID);
				if (textRuler == null)
				{
					return null;
				}
				tabStops = textRuler.GetTabStops();
			}
			//.stream().map(HSLFTabStopDecorator::new).collect(Collectors.toList());
			return tabStops as List<T>;
		}

		//@Override
		public void AddTabStops(double positionInPoints, TabStopType tabStopType)
		{
			HSLFTabStop ts = new HSLFTabStop(0, tabStopType);
			ts.SetPositionInPoints(positionInPoints);

			if (GetSheet() is HSLFSlideMaster)
			{
				Action<HSLFTabStopPropCollection> con = (tp) => tp.AddTabStop(ts);
				SetPropValInner(_paragraphStyle, HSLFTabStopPropCollection.NAME, (Action<TextProp>)con);
			}
			else
			{
				RecordContainer cont = _headerAtom.GetParentRecord();
				TextRulerAtom textRuler = (TextRulerAtom)cont.FindFirstOfType(RecordTypes.TextRulerAtom.typeID);
				if (textRuler == null)
				{
					textRuler = TextRulerAtom.GetParagraphInstance();
					cont.AppendChildRecord(textRuler);
				}
				textRuler.GetTabStops().Add(ts);
			}
		}

		//@Override
		public void ClearTabStops()
		{
			if (GetSheet() is HSLFSlideMaster)
			{
				SetPropValInner(_paragraphStyle, HSLFTabStopPropCollection.NAME, null);
			}
			else
			{
				RecordContainer cont = _headerAtom.GetParentRecord();
				TextRulerAtom textRuler = (TextRulerAtom)cont.FindFirstOfType(RecordTypes.TextRulerAtom.typeID);
				if (textRuler == null)
				{
					return;
				}
				textRuler.GetTabStops().Clear();
			}
		}

		private Double GetPctOrPoints(String propName)
		{
			TextProp tp = GetPropVal<TextProp>(_paragraphStyle, propName);
			if (tp == null)
			{
				return 0;
			}
			int val = tp.GetValue();
			return (val < 0) ? Units.MasterToPoints(val) : val;
		}

		private void SetPctOrPoints(String propName, Double dval)
		{
			int ival = 0;
			if (dval != null)
			{
				ival = (dval < 0) ? Units.PointsToMaster(dval) : (int)dval;
			}
			SetParagraphTextPropVal(propName, ival);
		}

		private bool GetFlag(int index)
		{
			BitMaskTextProp tp = GetPropVal<BitMaskTextProp>(_paragraphStyle, ParagraphFlagsTextProp.NAME);
			return tp != null && tp.GetSubValue(index);
		}

		private void SetFlag(int index, bool value)
		{
			BitMaskTextProp tp = _paragraphStyle.AddWithName<BitMaskTextProp>(ParagraphFlagsTextProp.NAME);
			tp.SetSubValue(value, index);
			SetDirty();
		}

		/**
		 * Fetch the value of the given Paragraph related TextProp. Returns null if
		 * that TextProp isn't present. If the TextProp isn't present, the value
		 * from the appropriate Master Sheet will apply.
		 *
		 * The propName can be a comma-separated list, in case multiple equivalent values
		 * are queried
		 */
		public T GetPropVal<T>(TextPropCollection props, String propName) where T : TextProp
		{
			String[] propNames = propName.Split(',');
			foreach (String pn in propNames)
			{
				T prop = props.FindByName<T>(pn);
				if (IsValidProp(prop))
				{
					return prop;
				}
			}

			return GetMasterPropVal<T>(props, propName);
		}

		private T GetMasterPropVal<T>(TextPropCollection props, String propName) where T : TextProp
		{
			bool isChar = props.GetTextPropType() == TextPropType.character;

			// check if we can delegate to master for the property
			if (!isChar)
			{
				BitMaskTextProp maskProp = props.FindByName<BitMaskTextProp>(ParagraphFlagsTextProp.NAME);
				bool hardAttribute = (maskProp != null && maskProp.GetValue() == 0);
				if (hardAttribute)
				{
					return null;
				}
			}

			String[] propNames = propName.Split(',');
			HSLFSheet sheet = GetSheet();
			int txtype = GetRunType();
			HSLFMasterSheet master;
			if (sheet is HSLFMasterSheet)
			{
				master = (HSLFMasterSheet)sheet;
			}
			else
			{
				master = sheet.GetMasterSheet();
				if (master == null)
				{
					//LOG.atWarn().log("MasterSheet is not available");
					return null;
				}
			}

			foreach (String pn in propNames)
			{
				TextPropCollection masterProps = master.GetPropCollection(txtype, GetIndentLevel(), pn, isChar);
				if (masterProps != null)
				{
					T prop = masterProps.FindByName<T>(pn);
					if (IsValidProp(prop))
					{
						return prop;
					}
				}
			}

			return null;
		}

		private static bool IsValidProp(TextProp prop)
		{
			// Font properties (maybe other too???) can have an index of -1
			// so we check the master for this font index then
			return prop != null && (!prop.GetName().Contains("font") || prop.GetValue() != -1);
		}

		/**
		 * Returns the named TextProp, either by fetching it (if it exists) or
		 * adding it (if it didn't)
		 *
		 * @param props the TextPropCollection to fetch from / add into
		 * @param name the name of the TextProp to fetch/add
		 * @param val the value, null if unset
		 */
		//public void SetPropVal(TextPropCollection props, String name, int val)
		//{
		//	Action<TextProp> act = val == 0 ? () => null : (value) => { TextProp tp = new TextProp(); tp.SetValue(val); return tp; };
		//	SetPropValInner(props, name, act);
		//}

		private void SetPropValInner(TextPropCollection props, String name, Action<TextProp> handler)
		{
			bool isChar = props.GetTextPropType() == TextPropType.character;

			TextPropCollection pc;
			if (_sheet is HSLFMasterSheet)
			{
				pc = ((HSLFMasterSheet)_sheet).GetPropCollection(GetRunType(), GetIndentLevel(), "*", isChar);
				if (pc == null)
				{
					throw new HSLFException("Master text property collection can't be determined.");
				}
			}
			else
			{
				pc = props;
			}

			if (handler == null)
			{
				pc.RemoveByName<TextProp>(name);
			}
			else
			{
				// Fetch / Add the TextProp
				handler(pc.AddWithName<TextProp>(name));
			}
			SetDirty();
		}


		/**
		 * Check and add linebreaks to text runs leading other paragraphs
		 */
		protected static void FixLineEndings(List<HSLFTextParagraph> paragraphs)
		{
			HSLFTextRun lastRun = null;
			foreach (HSLFTextParagraph p in paragraphs)
			{
				if (lastRun != null && !lastRun.GetRawText().EndsWith("\r"))
				{
					lastRun.SetText(lastRun.GetRawText() + "\r");
				}
				List<HSLFTextRun> ltr = p.GetTextRuns();
				if (ltr.Count == 0)
				{
					throw new HSLFException("paragraph without textruns found");
				}
				lastRun = ltr.ElementAt(ltr.Count - 1);
				//assert(lastRun.getRawText() != null);
			}
		}

		/**
		 * Search for a StyleTextPropAtom is for this text header (list of paragraphs)
		 *
		 * @param header the header
		 * @param textLen the length of the rawtext, or -1 if the length is not known
		 */
		private static StyleTextPropAtom FindStyleAtomPresent(TextHeaderAtom header, int textLen)
		{
			bool afterHeader = false;
			StyleTextPropAtom style = null;
			foreach (Record.Record record in header.GetParentRecord().GetChildRecords())
			{
				long rt = record.GetRecordType();
				if (afterHeader && rt == RecordTypes.TextHeaderAtom.typeID)
				{
					// already on the next header, quit searching
					break;
				}
				afterHeader |= (header == record);
				if (afterHeader && rt == RecordTypes.StyleTextPropAtom.typeID)
				{
					// found it
					style = (StyleTextPropAtom)record;
				}
			}

			if (style == null)
			{
				//LOG.atInfo().log("styles atom doesn't exist. Creating dummy record for later saving.");
				style = new StyleTextPropAtom((textLen < 0) ? 1 : textLen);
			}
			else
			{
				if (textLen >= 0)
				{
					style.SetParentTextSize(textLen);
				}
			}

			return style;
		}

		/**
		 * Saves the modified paragraphs/textrun to the records.
		 * Also updates the styles to the correct text length.
		 */
		protected static void StoreText(List<HSLFTextParagraph> paragraphs)
		{
			FixLineEndings(paragraphs);
			UpdateTextAtom(paragraphs);
			UpdateStyles(paragraphs);
			UpdateHyperlinks(paragraphs);
			RefreshRecords(paragraphs);

			foreach (HSLFTextParagraph p in paragraphs)
			{
				p._dirty = false;
			}
		}

		/**
		 * Set the correct text atom depending on the multibyte usage
		 */
		private static void UpdateTextAtom(List<HSLFTextParagraph> paragraphs)
		{
			String rawText = ToInternalString(GetRawText(paragraphs));

			// Will it fit in a 8 bit atom?
			bool isUnicode = StringUtil.HasMultibyte(rawText);
			// isUnicode = true;

			TextHeaderAtom headerAtom = paragraphs.ElementAt(0)._headerAtom;
			TextBytesAtom byteAtom = paragraphs.ElementAt(0)._byteAtom;
			TextCharsAtom charAtom = paragraphs.ElementAt(0)._charAtom;
			StyleTextPropAtom styleAtom = FindStyleAtomPresent(headerAtom, rawText.Length);

			// Store in the appropriate record
			Record.Record oldRecord = null, newRecord;
			if (isUnicode)
			{
				if (byteAtom != null || charAtom == null)
				{
					oldRecord = byteAtom;
					charAtom = new TextCharsAtom();
				}
				newRecord = charAtom;
				charAtom.SetText(rawText);
			}
			else
			{
				if (charAtom != null || byteAtom == null)
				{
					oldRecord = charAtom;
					byteAtom = new TextBytesAtom();
				}
				newRecord = byteAtom;
				byte[] byteText = new byte[rawText.Length];
				StringUtil.PutCompressedUnicode(rawText, byteText, 0);
				byteAtom.SetText(byteText);
			}

			RecordContainer _txtbox = headerAtom.GetParentRecord();
			Record.Record[] cr = _txtbox.GetChildRecords();
			int /* headerIdx = -1, */ textIdx = -1, styleIdx = -1;
			for (int i = 0; i < cr.Length; i++)
			{
				Record.Record r = cr[i];
				if (r == headerAtom)
				{
					// headerIdx = i;
				}
				else if (r == oldRecord || r == newRecord)
				{
					textIdx = i;
				}
				else if (r == styleAtom)
				{
					styleIdx = i;
				}
			}

			if (textIdx == -1)
			{
				// the old record was never registered, ignore it
				_txtbox.AddChildAfter(newRecord, headerAtom);
				// textIdx = headerIdx + 1;
			}
			else
			{
				// swap not appropriated records - noop if unchanged
				cr[textIdx] = newRecord;
			}

			if (styleIdx == -1)
			{
				// Add the new StyleTextPropAtom after the TextCharsAtom / TextBytesAtom
				_txtbox.AddChildAfter(styleAtom, newRecord);
			}

			foreach (HSLFTextParagraph p in paragraphs)
			{
				if (newRecord == byteAtom)
				{
					p._byteAtom = byteAtom;
					p._charAtom = null;
				}
				else
				{
					p._byteAtom = null;
					p._charAtom = charAtom;
				}
			}

		}

		/**
		 * Update paragraph and character styles - merges them when subsequential styles match
		 */
		private static void UpdateStyles(List<HSLFTextParagraph> paragraphs)
		{
			String rawText = ToInternalString(GetRawText(paragraphs));
			TextHeaderAtom headerAtom = paragraphs.ElementAt(0)._headerAtom;
			StyleTextPropAtom styleAtom = FindStyleAtomPresent(headerAtom, rawText.Length);

			// Update the text length for its Paragraph and Character stylings
			// * reset the length, to the new string's length
			// * add on +1 if the last block

			styleAtom.ClearStyles();

			TextPropCollection lastPTPC = null, lastRTPC = null, ptpc = null, rtpc = null;
			foreach (HSLFTextParagraph para in paragraphs)
			{
				ptpc = para.GetParagraphStyle();
				ptpc.UpdateTextSize(0);
				if (!ptpc.Equals(lastPTPC))
				{
					lastPTPC = ptpc.Copy();
					lastPTPC.UpdateTextSize(0);
					styleAtom.AddParagraphTextPropCollection(lastPTPC);
				}
				foreach (HSLFTextRun tr in para.GetTextRuns())
				{
					rtpc = tr.GetCharacterStyle();
					rtpc.UpdateTextSize(0);
					if (!rtpc.Equals(lastRTPC))
					{
						lastRTPC = rtpc.Copy();
						lastRTPC.UpdateTextSize(0);
						styleAtom.AddCharacterTextPropCollection(lastRTPC);
					}
					int len = tr.GetLength();
					ptpc.UpdateTextSize(ptpc.GetCharactersCovered() + len);
					rtpc.UpdateTextSize(len);
					lastPTPC.UpdateTextSize(lastPTPC.GetCharactersCovered() + len);
					lastRTPC.UpdateTextSize(lastRTPC.GetCharactersCovered() + len);
				}
			}

			if (lastPTPC == null || lastRTPC == null || ptpc == null || rtpc == null)
			{ // NOSONAR
				throw new HSLFException("Not all TextPropCollection could be determined.");
			}

			ptpc.UpdateTextSize(ptpc.GetCharactersCovered() + 1);
			rtpc.UpdateTextSize(rtpc.GetCharactersCovered() + 1);
			lastPTPC.UpdateTextSize(lastPTPC.GetCharactersCovered() + 1);
			lastRTPC.UpdateTextSize(lastRTPC.GetCharactersCovered() + 1);

			// If TextSpecInfoAtom is present, we must update the text size in it,
			// otherwise the ppt will be corrupted
			foreach (Record.Record r in paragraphs.ElementAt(0).GetRecords())
			{
				if (r is TextSpecInfoAtom)
				{
					((TextSpecInfoAtom)r).SetParentSize(rawText.Length + 1);
					break;
				}
			}
		}

		private static void UpdateHyperlinks(List<HSLFTextParagraph> paragraphs)
		{
			TextHeaderAtom headerAtom = paragraphs.ElementAt(0)._headerAtom;
			RecordContainer _txtbox = headerAtom.GetParentRecord();
			// remove existing hyperlink records
			foreach (Record.Record r in _txtbox.GetChildRecords())
			{
				if (r is InteractiveInfo || r is TxInteractiveInfoAtom)
				{
					_txtbox.RemoveChild(r);
				}
			}
			// now go through all the textruns and check for hyperlinks
			HSLFHyperlink lastLink = null;
			foreach (HSLFTextParagraph para in paragraphs)
			{
				foreach (HSLFTextRun run in para)
				{
					HSLFHyperlink thisLink = run.GetHyperlink();
					if (thisLink != null && thisLink == lastLink)
					{
						// the hyperlink extends over this text run, increase its length
						// TODO: the text run might be longer than the hyperlink
						thisLink.SetEndIndex(thisLink.GetEndIndex() + run.GetLength());
					}
					else
					{
						if (lastLink != null)
						{
							InteractiveInfo info = lastLink.GetInfo();
							TxInteractiveInfoAtom txinfo = lastLink.GetTextRunInfo();
							if (info != null && txinfo != null)
							{
								_txtbox.AppendChildRecord(info);
								_txtbox.AppendChildRecord(txinfo);
							}
						}
					}
					lastLink = thisLink;
				}
			}

			if (lastLink != null)
			{
				InteractiveInfo info = lastLink.GetInfo();
				TxInteractiveInfoAtom txinfo = lastLink.GetTextRunInfo();
				if (info != null && txinfo != null)
				{
					_txtbox.AppendChildRecord(info);
					_txtbox.AppendChildRecord(txinfo);
				}
			}
		}
		/**
		 * Writes the textbox records back to the document record
		 */
		private static void RefreshRecords(List<HSLFTextParagraph> paragraphs)
		{
			TextHeaderAtom headerAtom = paragraphs.ElementAt(0)._headerAtom;
			RecordContainer _txtbox = headerAtom.GetParentRecord();
			if (_txtbox is EscherTextboxWrapper)
			{
				try
				{
					_txtbox.WriteOut(null);
				}
				catch (IOException e)
				{
					throw new HSLFException("failed dummy write", e);
				}
			}
		}

		/**
		 * Adds the supplied text onto the end of the TextParagraphs,
		 * creating a new RichTextRun for it to sit in.
		 *
		 * @param text the text string used by this object.
		 */
		public static HSLFTextRun AppendText(List<HSLFTextParagraph> paragraphs, String text, bool newParagraph)
		{
			text = ToInternalString(text);

			// check paragraphs
			//assert(!paragraphs.isEmpty() && !paragraphs.get(0).getTextRuns().isEmpty());

			HSLFTextParagraph htp = paragraphs.ElementAt(paragraphs.Count - 1);
			HSLFTextRun htr = htp.GetTextRuns().ElementAt(htp.GetTextRuns().Count - 1);

			bool addParagraph = newParagraph;
			foreach (string rawText in text.Split("(?<=\r)".ToArray()))
			{
				// special case, if last text paragraph or run is empty, we will reuse it
				bool lastRunEmpty = (htr.GetLength() == 0);
				bool lastParaEmpty = lastRunEmpty && (htp.GetTextRuns().Count == 1);

				if (addParagraph && !lastParaEmpty)
				{
					TextPropCollection tpc = htp.GetParagraphStyle();
					HSLFTextParagraph prevHtp = htp;
					htp = new HSLFTextParagraph(htp._headerAtom, htp._byteAtom, htp._charAtom, paragraphs);
					htp.SetParagraphStyle(tpc.Copy());
					htp.SetParentShape(prevHtp.GetParentShape());
					htp.SetShapeId(prevHtp.GetShapeId());
					//htp.SupplySheet(prevHtp.GetSheet());
					paragraphs.Add(htp);
				}
				addParagraph = true;

				if (!lastRunEmpty)
				{
					TextPropCollection tpc = htr.GetCharacterStyle();
					htr = new HSLFTextRun(htp);
					htr.SetCharacterStyle(tpc.Copy());
					htp.AddTextRun(htr);
				}
				htr.SetText(rawText);
			}

			StoreText(paragraphs);

			return htr;
		}

		/**
		 * Sets (overwrites) the current text.
		 * Uses the properties of the first paragraph / textrun
		 *
		 * @param text the text string used by this object.
		 */
		//public static HSLFTextRun SetText(List<HSLFTextParagraph> paragraphs, String text)
		//{
		//	// check paragraphs
		//	if (!(paragraphs.Count == 0) && !(paragraphs.ElementAt(0).GetTextRuns().Count == 0))
		//	{

		//		Iterator<HSLFTextParagraph> paraIter = paragraphs.iterator();
		//		HSLFTextParagraph htp = paraIter.hasNext() ? paraIter.next() : null; // keep first
		//		if (htp != null)
		//		{
		//			while (paraIter.hasNext())
		//			{
		//				paraIter.next();
		//				paraIter.remove();
		//			}

		//			Iterator<HSLFTextRun> runIter = htp.GetTextRuns().iterator();
		//			if (runIter.hasNext())
		//			{
		//				HSLFTextRun htr = runIter.next();
		//				htr.SetText("");
		//				while (runIter.hasNext())
		//				{
		//					runIter.next();
		//					runIter.remove();
		//				}
		//			}
		//			else
		//			{
		//				HSLFTextRun trun = new HSLFTextRun(htp);
		//				htp.AddTextRun(trun);
		//			}
		//		}
		//	}

		//	return AppendText(paragraphs, text, false);
		//}

		public static String GetText(List<HSLFTextParagraph> paragraphs)
		{
			String rawText = String.Empty;
			if (!(paragraphs.Count == 0))
			{
				rawText = GetRawText(paragraphs);
			}
			return ToExternalString(rawText, paragraphs.ElementAt(0).GetRunType());
		}

		public static String GetRawText(List<HSLFTextParagraph> paragraphs)
		{
			StringBuilder sb = new StringBuilder();
			foreach (HSLFTextParagraph p in paragraphs)
			{
				foreach (HSLFTextRun r in p.GetTextRuns())
				{
					sb.Append(r.GetRawText());
				}
			}
			return sb.ToString();
		}

		//@Override
		public override string ToString()
		{
			StringBuilder sb = new StringBuilder();
			foreach (HSLFTextRun r in GetTextRuns())
			{
				sb.Append(r.GetRawText());
			}
			return ToExternalString(sb.ToString(), GetRunType());
		}

		/**
		 * Returns a new string with line breaks converted into internal ppt
		 * representation
		 */
		public static String ToInternalString(String s)
		{
			return s.Replace("\\r?\\n", "\r");
		}

		/**
		 * Converts raw text from the text paragraphs to a formatted string,
		 * i.e. it converts certain control characters used in the raw txt
		 *
		 * @param rawText the raw text
		 * @param runType the run type of the shape, paragraph or headerAtom.
		 *        use -1 if unknown
		 * @return the formatted string
		 */
		public static String ToExternalString(String rawText, int runType)
		{
			// PowerPoint seems to store files with \r as the line break
			// The messes things up on everything but a Mac, so translate
			// them to \n
			String text = rawText.Replace('\r', '\n');

			// 0xB acts like carriage return in page titles and like blank in the others
			char repl = (runType == -1 ||
				runType == (int)TextPlaceholderEnum.TITLE ||
				runType == (int)TextPlaceholderEnum.CENTER_TITLE) ? '\n' : ' ';

			return text.Replace('\u000b', repl);
		}

		/**
		 * For a given PPDrawing, grab all the TextRuns
		 */
		public static List<List<HSLFTextParagraph>> FindTextParagraphs(PPDrawing ppdrawing, HSLFSheet sheet)
		{
			if (ppdrawing == null)
			{
				throw new InvalidOperationException("Did not receive a valid drawing for sheet " + sheet._getSheetNumber());
			}

			List<List<HSLFTextParagraph>> runsV = new List<List<HSLFTextParagraph>>();
			foreach (EscherTextboxWrapper wrapper in ppdrawing.GetTextboxWrappers())
			{
				List<HSLFTextParagraph> p = FindTextParagraphs(wrapper, sheet);
				if (p != null)
				{
					runsV.Add(p);
				}
			}
			return runsV;
		}

		/**
		 * Scans through the supplied record array, looking for
		 * a TextHeaderAtom followed by one of a TextBytesAtom or
		 * a TextCharsAtom. Builds up TextRuns from these
		 *
		 * @param wrapper an EscherTextboxWrapper
		 */
		public static List<HSLFTextParagraph> FindTextParagraphs(EscherTextboxWrapper wrapper, HSLFSheet sheet)
		{
			// propagate parents to parent-aware records
			RecordContainer.HandleParentAwareRecords(wrapper);
			int shapeId = wrapper.getShapeId();
			List<HSLFTextParagraph> rv = null;

			OutlineTextRefAtom ota = (OutlineTextRefAtom)wrapper.FindFirstOfType(RecordTypes.OutlineTextRefAtom.typeID);
			if (ota != null)
			{
				// if we are based on an outline, there are no further records to be parsed from the wrapper
				if (sheet == null)
				{
					throw new HSLFException("Outline atom reference can't be solved without a sheet record");
				}

				List<List<HSLFTextParagraph>> sheetRuns = sheet.GetTextParagraphs();
				if (sheetRuns != null)
				{

					int idx = ota.getTextIndex();
					foreach (List<HSLFTextParagraph> r in sheetRuns)
					{
						if (r.Count==0)
						{
							continue;
						}
						int ridx = r.ElementAt(0).GetIndex();
						if (ridx > idx)
						{
							break;
						}
						if (ridx == idx)
						{
							if (rv == null)
							{
								rv = r;
							}
							else
							{
								// create a new container
								// TODO: ... is this case really happening?
								rv = new List<HSLFTextParagraph>();
								rv.AddRange(r);
							}
						}
					}
					if (rv == null || rv.Count==0)
					{
						//LOG.atWarn().log("text run not found for OutlineTextRefAtom.TextIndex={}", box(idx));
					}
				}
			}
			else
			{
				if (sheet != null)
				{
					// check sheet runs first, so we get exactly the same paragraph list
					List<List<HSLFTextParagraph>> sheetRuns = sheet.GetTextParagraphs();
					if (sheetRuns != null)
					{
						foreach (List<HSLFTextParagraph> paras in sheetRuns)
						{
							if (!(paras.Count==0) && paras.ElementAt(0)._headerAtom.GetParentRecord() == wrapper)
							{
								rv = paras;
								break;
							}
						}
					}
				}

				if (rv == null)
				{
					// if we haven't found the wrapper in the sheet runs, create a new paragraph list from its record
					List<List<HSLFTextParagraph>> rvl = FindTextParagraphs(wrapper.GetChildRecords());
					switch (rvl.Count)
					{
						case 0: break; // nothing found
						case 1: rv = rvl.ElementAt(0); break; // normal case
						default:
							throw new HSLFException("TextBox contains more than one list of paragraphs.");
					}
				}
			}

			if (rv != null)
			{
				StyleTextProp9Atom styleTextProp9Atom = wrapper.getStyleTextProp9Atom();

				foreach (HSLFTextParagraph htp in rv)
				{
					htp.SetShapeId(shapeId);
					htp.SetStyleTextProp9Atom(styleTextProp9Atom);
				}
			}
			return rv;
		}

		/**
		 * Scans through the supplied record array, looking for
		 * a TextHeaderAtom followed by one of a TextBytesAtom or
		 * a TextCharsAtom. Builds up TextRuns from these
		 *
		 * @param records the records to build from
		 */
		protected static List<List<HSLFTextParagraph>> FindTextParagraphs(Record.Record[] records)
		{
			List<List<HSLFTextParagraph>> paragraphCollection = new List<List<HSLFTextParagraph>>();

			int[] recordIdx = { 0 };

			for (int slwtIndex = 0; recordIdx[0] < records.Length; slwtIndex++)
			{
				TextHeaderAtom header = null;
				TextBytesAtom tbytes = null;
				TextCharsAtom tchars = null;
				TextRulerAtom ruler = null;
				MasterTextPropAtom indents = null;

				foreach (Record.Record r in GetRecords(records, recordIdx, null))
				{
					long rt = r.GetRecordType();
					if (RecordTypes.TextHeaderAtom.typeID == rt)
					{
						header = (TextHeaderAtom)r;
					}
					else if (RecordTypes.TextBytesAtom.typeID == rt)
					{
						tbytes = (TextBytesAtom)r;
					}
					else if (RecordTypes.TextCharsAtom.typeID == rt)
					{
						tchars = (TextCharsAtom)r;
					}
					else if (RecordTypes.TextRulerAtom.typeID == rt)
					{
						ruler = (TextRulerAtom)r;
					}
					else if (RecordTypes.MasterTextPropAtom.typeID == rt)
					{
						indents = (MasterTextPropAtom)r;
					}
					// don't search for RecordTypes.StyleTextPropAtom.typeID here ... see findStyleAtomPresent below
				}

				if (header == null)
				{
					break;
				}

				if (header.GetParentRecord() is SlideListWithText)
				{
					// runs found in PPDrawing are not linked with SlideListWithTexts
					header.SetIndex(slwtIndex);
				}

				if (tbytes == null && tchars == null)
				{
					tbytes = new TextBytesAtom();
					// don't add record yet - set it in storeText
					//LOG.atInfo().log("bytes nor chars atom doesn't exist. Creating dummy record for later saving.");
				}

				String rawText = (tchars != null) ? tchars.GetText() : tbytes.GetText();
				StyleTextPropAtom styles = FindStyleAtomPresent(header, rawText.Length);

				List<HSLFTextParagraph> paragraphs = new List<HSLFTextParagraph>();
				paragraphCollection.Add(paragraphs);

				// split, but keep delimiter
				foreach (String para in rawText.Split("(?<=\r)".ToArray()))
				{
					HSLFTextParagraph tpara = new HSLFTextParagraph(header, tbytes, tchars, paragraphs);
					paragraphs.Add(tpara);
					tpara._ruler = ruler;
					tpara.GetParagraphStyle().UpdateTextSize(para.Length);

					HSLFTextRun trun = new HSLFTextRun(tpara);
					tpara.AddTextRun(trun);
					trun.SetText(para);
				}

				ApplyCharacterStyles(paragraphs, styles.GetCharacterStyles());
				ApplyParagraphStyles(paragraphs, styles.GetParagraphStyles());
				if (indents != null)
				{
					ApplyParagraphIndents(paragraphs, indents.getIndents());
				}
			}

			if (paragraphCollection.Count == 0)
			{
				//LOG.atDebug().log("No text records found.");
			}

			return paragraphCollection;
		}

		protected static void ApplyHyperlinks(List<HSLFTextParagraph> paragraphs)
		{
			List<HSLFHyperlink> links = HSLFHyperlink.Find(paragraphs);

			foreach (HSLFHyperlink h in links)
			{
				int csIdx = 0;
				foreach (HSLFTextParagraph p in paragraphs)
				{
					if (csIdx > h.GetEndIndex())
					{
						break;
					}
					List<HSLFTextRun> runs = p.GetTextRuns();
					for (int rlen, rIdx = 0; rIdx < runs.Count; csIdx += rlen, rIdx++)
					{
						HSLFTextRun run = runs.ElementAt(rIdx);
						rlen = run.GetLength();
						if (csIdx < h.GetEndIndex() && h.GetStartIndex() < csIdx + rlen)
						{
							String rawText = run.GetRawText();
							int startIdx = h.GetStartIndex() - csIdx;
							if (startIdx > 0)
							{
								// hyperlink starts within current textrun
								HSLFTextRun newRun = new HSLFTextRun(p);
								newRun.SetCharacterStyle(run.GetCharacterStyle());
								newRun.SetText(rawText.Substring(startIdx));
								run.SetText(rawText.Substring(0, startIdx));
								runs.Insert(rIdx + 1, newRun);
								rlen = startIdx;
								continue;
							}
							int endIdx = Math.Min(rlen, h.GetEndIndex() - h.GetStartIndex());
							if (endIdx < rlen)
							{
								// hyperlink ends before end of current textrun
								HSLFTextRun newRun = new HSLFTextRun(p);
								newRun.SetCharacterStyle(run.GetCharacterStyle());
								newRun.SetText(rawText.Substring(0, endIdx));
								run.SetText(rawText.Substring(endIdx));
								runs.Insert(rIdx, newRun);
								rlen = endIdx;
								run = newRun;
							}
							run.SetHyperlink(h);
						}
					}
				}
			}
		}

		protected static void ApplyCharacterStyles(List<HSLFTextParagraph> paragraphs, List<TextPropCollection> charStyles)
		{
			int paraIdx = 0, runIdx = 0;
			HSLFTextRun trun;

			for (int csIdx = 0; csIdx < charStyles.Count; csIdx++)
			{
				TextPropCollection p = charStyles.ElementAt(csIdx);
				int ccStyle = p.GetCharactersCovered();
				if (ccStyle > MAX_NUMBER_OF_STYLES)
				{
					throw new InvalidOperationException("Cannot process more than " + MAX_NUMBER_OF_STYLES + " styles, but had paragraph with " + ccStyle);
				}
				for (int ccRun = 0; ccRun < ccStyle;)
				{
					HSLFTextParagraph para = paragraphs.ElementAt(paraIdx);
					List<HSLFTextRun> runs = para.GetTextRuns();
					trun = runs.ElementAt(runIdx);
					int len = trun.GetLength();

					if (ccRun + len <= ccStyle)
					{
						ccRun += len;
					}
					else
					{
						String text = trun.GetRawText();
						trun.SetText(text.Substring(0, ccStyle - ccRun));

						HSLFTextRun nextRun = new HSLFTextRun(para);
						nextRun.SetText(text.Substring(ccStyle - ccRun));
						runs.Insert(runIdx + 1, nextRun);

						ccRun += ccStyle - ccRun;
					}

					trun.SetCharacterStyle(p);

					if (paraIdx == paragraphs.Count - 1 && runIdx == runs.Count - 1)
					{
						if (csIdx < charStyles.Count - 1)
						{
							// special case, empty trailing text run
							HSLFTextRun nextRun = new HSLFTextRun(para);
							nextRun.SetText("");
							runs.Add(nextRun);
							ccRun++;
						}
						else
						{
							// need to add +1 to the last run of the last paragraph
							trun.GetCharacterStyle().UpdateTextSize(trun.GetLength() + 1);
							ccRun++;
						}
					}

					// need to compare it again, in case a run has been added after
					if (++runIdx == runs.Count)
					{
						paraIdx++;
						runIdx = 0;
					}
				}
			}
		}

		protected static void ApplyParagraphStyles(List<HSLFTextParagraph> paragraphs, List<TextPropCollection> paraStyles)
		{
			int paraIdx = 0;
			foreach (TextPropCollection p in paraStyles)
			{
				for (int ccPara = 0, ccStyle = p.GetCharactersCovered(); ccPara < ccStyle; paraIdx++)
				{
					if (paraIdx >= paragraphs.Count)
					{
						return;
					}
					HSLFTextParagraph htp = paragraphs.ElementAt(paraIdx);
					TextPropCollection pCopy = p.Copy();
					htp.SetParagraphStyle(pCopy);
					int len = 0;
					foreach (HSLFTextRun trun in htp.GetTextRuns())
					{
						len += trun.GetLength();
					}
					if (paraIdx == paragraphs.Count - 1)
					{
						len++;
					}
					pCopy.UpdateTextSize(len);
					ccPara += len;
				}
			}
		}

		protected static void ApplyParagraphIndents(List<HSLFTextParagraph> paragraphs, List<IndentProp> paraStyles)
		{
			int paraIdx = 0;
			foreach (IndentProp p in paraStyles)
			{
				for (int ccPara = 0, ccStyle = p.getCharactersCovered(); ccPara < ccStyle; paraIdx++)
				{
					if (paraIdx >= paragraphs.Count || ccPara >= ccStyle - 1)
					{
						return;
					}
					HSLFTextParagraph para = paragraphs.ElementAt(paraIdx);
					int len = 0;
					foreach (HSLFTextRun trun in para.GetTextRuns())
					{
						len += trun.GetLength();
					}
					para.SetIndentLevel(p.getIndentLevel());
					ccPara += len + 1;
				}
			}
		}

		public EscherTextboxWrapper GetTextboxWrapper()
		{
			return (EscherTextboxWrapper)_headerAtom.GetParentRecord();
		}
/**
		 * Sets the value of the given Paragraph TextProp, add if required
		 * @param propName The name of the Paragraph TextProp
		 * @param val The value to set for the TextProp
		 */
		public void SetParagraphTextPropVal(String propName, int val)
		{
			//SetPropVal(_paragraphStyle, propName, val);
			SetDirty();
		}

		/**
		 * marks this paragraph dirty, so its records will be renewed on save
		 */
		public void SetDirty()
		{
			_dirty = true;
		}
public bool IsDirty()
		{
			return _dirty;
		}

		/**
		 * Calculates the start index of the given text run
		 *
		 * @param textrun the text run to search for
		 * @return the start index with the paragraph collection or -1 if not found
		 */
		/* package */
		public int GetStartIdxOfTextRun(HSLFTextRun textrun)
		{
			int idx = 0;
			foreach (HSLFTextParagraph p in parentList)
			{
				foreach (HSLFTextRun r in p)
				{
					if (r == textrun)
					{
						return idx;
					}
					idx += r.GetLength();
				}
			}
			return -1;
		}

		/**
		 * @see RoundTripHFPlaceholder12
		 */
		//@Override
		public bool IsHeaderOrFooter()
		{
			HSLFTextShape s = GetParentShape();
			if (s == null)
			{
				return false;
			}
			Placeholder ph = s.getPlaceholder();
			if (ph == null)
			{
				return false;
			}
			switch (ph.nativeEnum)
			{
				case "DATETIME":
				case "SLIDE_NUMBER":
				case "FOOTER":
				case "HEADER":
					return true;
				default:
					return false;
			}
		}

		public List<TabStop> GetTabStops()
		{
			throw new NotImplementedException();
		}

		public void AddTabStops(double positionInPoints, TabStopTypeEnum tabStopType)
		{
			throw new NotImplementedException();
		}

		public IEnumerator<HSLFTextRun> GetEnumerator()
		{
			throw new NotImplementedException();
		}

		IEnumerator<HSLFTextRun> IEnumerable<HSLFTextRun>.GetEnumerator()
		{
			return _runs.GetEnumerator();
		}

		System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
		{
			return _runs.GetEnumerator();
		}
	}
}
