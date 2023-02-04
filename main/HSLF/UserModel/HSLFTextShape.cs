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

using NPOI.Common.UserModel;
using NPOI.DDF;
using NPOI.HSLF.Exceptions;
using NPOI.HSLF.Model;
using NPOI.HSLF.Record;
using NPOI.POIFS.Properties;
using NPOI.SL.UserModel;
using NPOI.Util;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace NPOI.HSLF.UserModel
{

    /**
     * A common superclass of all shapes that can hold text.
     */
    public abstract class HSLFTextShape : HSLFShape//, TextShape<HSLFShape, HSLFTextParagraph>
    {

        /**
         * How to anchor the text
         */
        private enum HSLFTextAnchorEnum
        {
            TOP = 0,
            MIDDLE = 1,
            BOTTOM = 2,
            TOP_CENTER = 3,
            MIDDLE_CENTER = 4,
            BOTTOM_CENTER = 5,
            TOP_BASELINE = 6,
            BOTTOM_BASELINE = 7,
            TOP_CENTER_BASELINE = 8,
            BOTTOM_CENTER_BASELINE = 9,
        }

        public class HSLFTextAnchor
        {
            public int nativeId;
            public VerticalAlignment vAlign;
            public bool centered;
            public bool baseline;

            public HSLFTextAnchor(int nativeId, VerticalAlignment vAlign, bool centered, bool baseline)
            {
                this.nativeId = nativeId;
                this.vAlign = vAlign;
                this.centered = centered;
                this.baseline = baseline;
            }
        }

        private static HSLFTextAnchor[] HSLFTextAnchors = new[]
        {
            new HSLFTextAnchor(0, VerticalAlignment.TOP,    false, false),
            new HSLFTextAnchor(1, VerticalAlignment.MIDDLE, false, false),
            new HSLFTextAnchor(2, VerticalAlignment.BOTTOM, false, false),
            new HSLFTextAnchor(3, VerticalAlignment.TOP,    true,  false),
            new HSLFTextAnchor(4, VerticalAlignment.MIDDLE, true,  false),
            new HSLFTextAnchor(5, VerticalAlignment.BOTTOM, true,  false),
            new HSLFTextAnchor(6, VerticalAlignment.TOP,    false, true),
            new HSLFTextAnchor(7, VerticalAlignment.BOTTOM, false, true),
            new HSLFTextAnchor(8, VerticalAlignment.TOP,    true,  true),
            new HSLFTextAnchor(9, VerticalAlignment.BOTTOM, true,  true)
    };
        static HSLFTextAnchor fromHSLFTextAnchorId(int nativeId)
        {
            if (0 <= nativeId && nativeId < HSLFTextAnchors.Length)
            {
                return HSLFTextAnchors[nativeId];
            }
            return null;
        }

        static HSLFTextAnchor fromHSLFTextAnchorEnum(HSLFTextAnchorEnum anchor)
        {
            return fromHSLFTextAnchorId((int)anchor);
        }


        /**
         * Specifies that a line of text will continue on subsequent lines instead
         * of extending into or beyond a margin.
         * Office Excel 2007, Excel 2010, PowerPoint 97, and PowerPoint 2010 read
         * and use this value properly but do not write it.
         */
        public static int WrapSquare = 0;
        /**
         * Specifies a wrapping rule that is equivalent to that of WrapSquare
         * Excel 97, Excel 2000, Excel 2002, and Office Excel 2003 use this value.
         * All other product versions listed at the beginning of this appendix ignore this value.
         */
        public static int WrapByPoints = 1;
        /**
         * Specifies that a line of text will extend into or beyond a margin instead
         * of continuing on subsequent lines.
         * Excel 97, Word 97, Excel 2000, Word 2000, Excel 2002,
         * and Office Excel 2003 do not use this value.
         */
        public static int WrapNone = 2;
        /**
         * Specifies a wrapping rule that is undefined and MUST be ignored.
         */
        public static int WrapTopBottom = 3;
        /**
         * Specifies a wrapping rule that is undefined and MUST be ignored.
         */
        public static int WrapThrough = 4;


        /**
         * TextRun object which holds actual text and format data
         */
        private List<HSLFTextParagraph> _paragraphs = new List<HSLFTextParagraph>();

        /**
         * Escher container which holds text attributes such as
         * TextHeaderAtom, TextBytesAtom or TextCharsAtom, StyleTextPropAtom etc.
         */
        private EscherTextboxWrapper _txtbox;

        /**
         * This setting is used for supporting a deprecated alignment
         *
         * @see <a href=""></a>
         */
        //    bool alignToBaseline = false;

        /**
         * Create a TextBox object and initialize it from the supplied Record container.
         *
         * @param escherRecord       {@code EscherSpContainer} container which holds information about this shape
         * @param parent    the parent of the shape
         */
        //  protected HSLFTextShape(EscherContainerRecord escherRecord, ShapeContainer<HSLFShape,HSLFTextParagraph> parent)
        //          : base(escherRecord, parent)
        //{

        //  }

        /**
         * Create a new TextBox. This constructor is used when a new shape is created.
         *
         * @param parent    the parent of this Shape. For example, if this text box is a cell
         * in a table then the parent is Table.
         */
        //public HSLFTextShape(ShapeContainer<HSLFShape,HSLFTextParagraph> parent)
        //        :base(null, parent){

        //    createSpContainer(parent is HSLFGroupShape);
        //}

        /**
         * Create a new TextBox. This constructor is used when a new shape is created.
         *
         */
        public HSLFTextShape(EscherContainerRecord escherRecord, ShapeContainer<HSLFShape, HSLFTextParagraph> parent)
		: base(escherRecord, parent)
		{

        }

        /**
         * Set default properties for the  TextRun.
         * Depending on the text and shape type the defaults are different:
         *   TextBox: align=left, valign=top
         *   AutoShape: align=center, valign=middle
         *
         */
        protected void setDefaultTextProperties(HSLFTextParagraph _txtrun)
        {

        }

        /**
         * When a textbox is added to  a sheet we need to tell upper-level
         * {@code PPDrawing} about it.
         *
         * @param sh the sheet we are adding to
         */
        //protected new void AfterInsert(HSLFSheet sh)
        //    {
        //    base.AfterInsert(sh);

        //    storeText();

        //    EscherTextboxWrapper thisTxtbox = getEscherTextboxWrapper();
        //    if(thisTxtbox != null){
        //        GetSpContainer().AddChildRecord(thisTxtbox.getEscherRecord());

        //        PPDrawing ppdrawing = sh.GetPPDrawing();
        //        ppdrawing.addTextboxWrapper(thisTxtbox);
        //        // Ensure the escher layer knows about the added records
        //        try {
        //            thisTxtbox.WriteOut(null);
        //        } catch (IOException e){
        //            throw new HSLFException(e);
        //        }
        //        bool isInitialAnchor = getAnchor().equals(new Rectangle2D.Double());
        //            bool isFilledTxt = "" != getText();
        //        if (isInitialAnchor && isFilledTxt) {
        //            resizeToFitText();
        //        }
        //    }
        //    foreach (HSLFTextParagraph htp in _paragraphs) {
        //        htp.setShapeId(GetShapeId());
        //    }
        //    sh.onAddTextShape(this);
        //}

        protected EscherTextboxWrapper getEscherTextboxWrapper()
        {
            if (_txtbox != null)
            {
                return _txtbox;
            }

            EscherTextboxRecord textRecord = (EscherTextboxRecord)GetEscherChild(EscherTextboxRecord.RECORD_ID);
            if (textRecord == null)
            {
                return null;
            }

            HSLFSheet sheet = GetSheet();
            if (sheet != null)
            {
                PPDrawing drawing = sheet.GetPPDrawing();
                if (drawing != null)
                {
                    EscherTextboxWrapper[] wrappers = drawing.GetTextboxWrappers();
                    if (wrappers != null)
                    {
                        foreach (EscherTextboxWrapper w in wrappers)
                        {
                            // check for object identity
                            if (textRecord == w.getEscherRecord())
                            {
                                _txtbox = w;
                                return _txtbox;
                            }
                        }
                    }
                }
            }

            _txtbox = new EscherTextboxWrapper(textRecord);
            return _txtbox;
        }

        private void createEmptyParagraph()
        {
            TextHeaderAtom tha = (TextHeaderAtom)_txtbox.FindFirstOfType(TextHeaderAtom._type);
            if (tha == null)
            {
                tha = new TextHeaderAtom();
                tha.SetParentRecord(_txtbox);
                _txtbox.AppendChildRecord(tha);
            }

            TextBytesAtom tba = (TextBytesAtom)_txtbox.FindFirstOfType(TextBytesAtom._type);
            TextCharsAtom tca = (TextCharsAtom)_txtbox.FindFirstOfType(TextCharsAtom._type);
            if (tba == null && tca == null)
            {
                tba = new TextBytesAtom();
                tba.SetText(new byte[0]);
                _txtbox.AppendChildRecord(tba);
            }

            String text = ((tba != null) ? tba.GetText() : tca.GetText());

            StyleTextPropAtom sta = (StyleTextPropAtom)_txtbox.FindFirstOfType(StyleTextPropAtom._type);
            TextPropCollection paraStyle = null, charStyle = null;
            if (sta == null)
            {
                int parSiz = text.Length;
                sta = new StyleTextPropAtom(parSiz + 1);
                if (_paragraphs.Count == 0)
                {
                    paraStyle = sta.AddParagraphTextPropCollection(parSiz + 1);
                    charStyle = sta.AddCharacterTextPropCollection(parSiz + 1);
                }
                else
                {
                    foreach (HSLFTextParagraph htp in _paragraphs)
                    {
                        int runsLen = 0;
                        foreach (HSLFTextRun htr in htp.GetTextRuns())
                        {
                            runsLen += htr.GetLength();
                            charStyle = sta.AddCharacterTextPropCollection(htr.GetLength());
                            htr.SetCharacterStyle(charStyle);
                        }
                        paraStyle = sta.AddParagraphTextPropCollection(runsLen);
                        htp.SetParagraphStyle(paraStyle);
                    }
                    //assert (paraStyle != null && charStyle != null);
                }
                _txtbox.AppendChildRecord(sta);
            }
            else
            {
                paraStyle = sta.GetParagraphStyles().ElementAt(0);
                charStyle = sta.GetCharacterStyles().ElementAt(0);
            }

            if (_paragraphs.Count == 0)
            {
                HSLFTextParagraph htp = new HSLFTextParagraph(tha, tba, tca, _paragraphs);
                htp.SetParagraphStyle(paraStyle);
                htp.SetParentShape(this);
                _paragraphs.Add(htp);

                HSLFTextRun htr = new HSLFTextRun(htp);
                htr.SetCharacterStyle(charStyle);
                htr.SetText(text);
                htp.AddTextRun(htr);
            }
        }

        //public override Rectangle2D resizeToFitText() {
        //    return resizeToFitText(null);
        //}

        //public override Rectangle2D resizeToFitText(Graphics2D graphics) {
        //    Rectangle2D anchor = getAnchor();
        //    if(anchor.getWidth() == 0.) {
        //        LOG.atWarn().log("Width of shape wasn't set. Defaulting to 200px");
        //        anchor.setRect(anchor.getX(), anchor.getY(), 200., anchor.getHeight());
        //        setAnchor(anchor);
        //    }
        //    double height = getTextHeight(graphics);
        //    height += 1; // add a pixel to compensate rounding errors

        //    Insets2D insets = getInsets();
        //    anchor.setRect(anchor.getX(), anchor.getY(), anchor.getWidth(), height+insets.top+insets.bottom);
        //    setAnchor(anchor);

        //    return anchor;
        //}

        /**
        * Returns the type of the text, from the TextHeaderAtom.
        * Possible values can be seen from TextHeaderAtom
        * @see TextHeaderAtom
        */
        public int getRunType()
        {
            getEscherTextboxWrapper();
            if (_txtbox == null)
            {
                return -1;
            }
            List<HSLFTextParagraph> paras = HSLFTextParagraph.FindTextParagraphs(_txtbox, GetSheet());
            return (paras.Count == 0 || paras.ElementAt(0).GetIndex() == -1) ? -1 : paras.ElementAt(0).GetRunType();
        }

        /**
        * Changes the type of the text. Values should be taken
        *  from TextHeaderAtom. No checking is done to ensure you
        *  set this to a valid value!
        * @see TextHeaderAtom
        */
        public void setRunType(int type)
        {
            getEscherTextboxWrapper();
            if (_txtbox == null)
            {
                return;
            }
            List<HSLFTextParagraph> paras = HSLFTextParagraph.FindTextParagraphs(_txtbox, GetSheet());
            if (!(paras.Count == 0))
            {
                paras.ElementAt(0).SetRunType(type);
            }
        }

        /**
         * Returns the type of vertical alignment for the text.
         * One of the <code>Anchor*</code> constants defined in this class.
         *
         * @return the type of alignment
         */
        ///* package */ HSLFTextAnchor getAlignment(){
        //    AbstractEscherOptRecord opt = GetEscherOptRecord();
        //    EscherSimpleProperty prop = GetEscherProperty(opt, EscherProperties.TEXT__ANCHORTEXT);
        //    HSLFTextAnchor align;
        //    if (prop == null){
        //        /**
        //         * If vertical alignment was not found in the shape properties then try to
        //         * fetch the master shape and search for the align property there.
        //         */
        //        int type = getRunType();
        //        HSLFSheet sh = GetSheet();
        //        HSLFMasterSheet master = (sh != null) ? sh.GetMasterSheet() : null;
        //        HSLFTextShape masterShape = (master != null) ? master.GetPlaceholderByTextType(type) : null;
        //        if (masterShape != null && type != (int)TextPlaceholderEnum.OTHER) {
        //            align = masterShape.getAlignment();
        //        } else {
        //                //not found in the master sheet. Use the hardcoded defaults.
        //                align = (TextPlaceholder.isTitle(type)) ? fromHSLFTextAnchorEnum(HSLFTextAnchorEnum.MIDDLE) : fromHSLFTextAnchorEnum(HSLFTextAnchor.TOP);
        //        }
        //    } else {
        //        align = fromHSLFTextAnchorEnum((HSLFTextAnchorEnum)prop.PropertyValue);
        //    }

        //    return (align == null) ?  HSLFTextAnchor.TOP : align;
        //}

        /**
         * Sets the type of alignment for the text.
         * One of the {@code Anchor*} constants defined in this class.
         *
         * @param isCentered horizontal centered?
         * @param vAlign vertical alignment
         * @param baseline aligned to baseline?
         */
        ///* package */ void setAlignment(bool isCentered, VerticalAlignment vAlign, bool baseline) {
        //    for (HSLFTextAnchor hta : HSLFTextAnchor.values()) {
        //        if (
        //            (hta.centered == (isCentered != null && isCentered)) &&
        //            (hta.vAlign == vAlign) &&
        //            (hta.baseline == null || hta.baseline == baseline)
        //        ) {
        //            setEscherProperty(EscherProperties.TEXT__ANCHORTEXT, hta.nativeId);
        //            break;
        //        }
        //    }
        //}

        /**
         * @return true, if vertical alignment is relative to baseline
         * this is only used for older versions less equals Office 2003
         */
        //public bool isAlignToBaseline() {
        //    return getAlignment().baseline;
        //}

        /**
         * Sets the vertical alignment relative to the baseline
         *
         * @param alignToBaseline if true, vertical alignment is relative to baseline
         */
        //public void setAlignToBaseline(bool alignToBaseline) {
        //    setAlignment(isHorizontalCentered(), getVerticalAlignment(), alignToBaseline);
        //}


        //public bool isHorizontalCentered() {
        //    return getAlignment().centered;
        //}


        //public void setHorizontalCentered(bool isCentered) {
        //    setAlignment(isCentered, getVerticalAlignment(), getAlignment().baseline);
        //}


        //public VerticalAlignment getVerticalAlignment() {
        //    return getAlignment().vAlign;
        //}


        //public void setVerticalAlignment(VerticalAlignment vAlign) {
        //    setAlignment(isHorizontalCentered(), vAlign, getAlignment().baseline);
        //}

        /**
         * Returns the distance (in points) between the bottom of the text frame
         * and the bottom of the inscribed rectangle of the shape that contains the text.
         * Default value is 1/20 inch.
         *
         * @return the botom margin
         */
        //public double getBottomInset(){
        //    return getInset(EscherProperties.TEXT__TEXTBOTTOM, .05);
        //}

        /**
         * Sets the botom margin.
         * @see #getBottomInset()
         *
         * @param margin    the bottom margin
         */
        //public void setBottomInset(double margin){
        //    setInset(EscherProperties.TEXT__TEXTBOTTOM, margin);
        //}

        /**
         *  Returns the distance (in points) between the left edge of the text frame
         *  and the left edge of the inscribed rectangle of the shape that contains
         *  the text.
         *  Default value is 1/10 inch.
         *
         * @return the left margin
         */
        //public double getLeftInset(){
        //    return getInset(EscherProperties.TEXT__TEXTLEFT, .1);
        //}

        /**
         * Sets the left margin.
         * @see #getLeftInset()
         *
         * @param margin    the left margin
         */
        //public void setLeftInset(double margin){
        //    setInset(EscherProperties.TEXT__TEXTLEFT, margin);
        //}

        /**
         *  Returns the distance (in points) between the right edge of the
         *  text frame and the right edge of the inscribed rectangle of the shape
         *  that contains the text.
         *  Default value is 1/10 inch.
         *
         * @return the right margin
         */
        //public double getRightInset(){
        //    return getInset(EscherProperties.TEXT__TEXTRIGHT, .1);
        //}

        /**
         * Sets the right margin.
         * @see #getRightInset()
         *
         * @param margin    the right margin
         */
        //public void setRightInset(double margin){
        //    setInset(EscherProperties.TEXT__TEXTRIGHT, margin);
        //}

        /**
        *  Returns the distance (in points) between the top of the text frame
        *  and the top of the inscribed rectangle of the shape that contains the text.
        *  Default value is 1/20 inch.
        *
        * @return the top margin
        */
        //public double getTopInset(){
        //    return getInset(EscherProperties.TEXT__TEXTTOP, .05);
        //}

        /**
          * Sets the top margin.
          * @see #getTopInset()
          *
          * @param margin    the top margin
          */
        //public void setTopInset(double margin){
        //    setInset(EscherProperties.TEXT__TEXTTOP, margin);
        //}

        /**
         * Returns the distance (in points) between the edge of the text frame
         * and the edge of the inscribed rectangle of the shape that contains the text.
         * Default value is 1/20 inch.
         *
         * @param type the type of the inset edge
         * @return the inset in points
         */
        //private double getInset(EscherProperties type, double defaultInch) {
        //    AbstractEscherOptRecord opt = GetEscherOptRecord();
        //    EscherSimpleProperty prop = GetEscherProperty(opt, type);
        //    int val = prop == null ? (int)(Units.ToEMU(Units.POINT_DPI)*defaultInch) : prop.getPropertyValue();
        //    return Units.ToPoints(val);
        //}

        /**
         * @param type the type of the inset edge
         * @param margin the inset in points
         */
        //private void setInset(EscherProperties type, double margin){
        //    SetEscherProperty(type, Units.ToEMU(margin));
        //}

        /**
         * Returns the value indicating word wrap.
         *
         * @return the value indicating word wrap.
         *  Must be one of the {@code Wrap*} constants defined in this class.
         *
         * @see <a href="https://msdn.microsoft.com/en-us/library/dd948168(v=office.12).aspx">MSOWRAPMODE</a>
         */
        public int getWordWrapEx()
        {
            AbstractEscherOptRecord opt = GetEscherOptRecord();
            EscherSimpleProperty prop = GetEscherProperty<EscherSimpleProperty>(opt, EscherProperties.TEXT__WRAPTEXT);
            return prop == null ? WrapSquare : prop.PropertyValue;
        }

        /**
         *  Specifies how the text should be wrapped
         *
         * @param wrap  the value indicating how the text should be wrapped.
         *  Must be one of the {@code Wrap*} constants defined in this class.
         */
        //public void setWordWrapEx(int wrap){
        //    setEscherProperty(EscherProperties.TEXT__WRAPTEXT, wrap);
        //}

        public bool getWordWrap()
        {
            int ww = getWordWrapEx();
            return (ww != WrapNone);
        }

        //public void setWordWrap(bool wrap) {
        //    setWordWrapEx(wrap ? WrapSquare : WrapNone);
        //}

        /**
         * @return id for the text.
         */
        public int getTextId()
        {
            AbstractEscherOptRecord opt = GetEscherOptRecord();
            EscherSimpleProperty prop = GetEscherProperty<EscherSimpleProperty>(opt, EscherProperties.TEXT__TEXTID);
            return prop == null ? 0 : prop.PropertyValue;
        }

        /**
         * Sets text ID
         *
         * @param id of the text
         */
        //public void setTextId(int id){
        //    SetEscherProperty(EscherProperties.TEXT__TEXTID, id);
        //}

        public List<HSLFTextParagraph> getTextParagraphs()
        {
            if (_paragraphs.Any())
            {
                return _paragraphs;
            }

            _txtbox = getEscherTextboxWrapper();
            if (_txtbox == null)
            {
                _txtbox = new EscherTextboxWrapper();
                createEmptyParagraph();
            }
            else
            {
                List<HSLFTextParagraph> pList = HSLFTextParagraph.FindTextParagraphs(_txtbox, GetSheet());
                if (pList == null)
                {
                    // there are actually TextBoxRecords without extra data - see #54722
                    createEmptyParagraph();
                }
                else
                {
                    _paragraphs = pList;
                }

                if (!_paragraphs.Any())
                {
                    //LOG.atWarn().log("TextRecord didn't contained any text lines");
                }
            }

            foreach (HSLFTextParagraph p in _paragraphs)
            {
                p.SetParentShape(this);
            }

            return _paragraphs;
        }


        //public override void setSheet(HSLFSheet sheet) {
        //    base.SetSheet(sheet);

        //    // Initialize _txtrun object.
        //    // (We can't do it in the constructor because the sheet
        //    //  is not assigned then, it's only built once we have
        //    //  all the records)
        //    List<HSLFTextParagraph> ltp = getTextParagraphs();
        //    HSLFTextParagraph.SupplySheet(ltp, sheet);
        //}

        /**
         * Return {@link OEPlaceholderAtom}, the atom that describes a placeholder.
         *
         * @return {@link OEPlaceholderAtom} or {@code null} if not found
         */
        public OEPlaceholderAtom getPlaceholderAtom()
        {
            return GetClientDataRecord<OEPlaceholderAtom>(RecordTypes.OEPlaceholderAtom.typeID);
        }

        /**
         * Return {@link RoundTripHFPlaceholder12}, the atom that describes a header/footer placeholder.
         * Compare the {@link RoundTripHFPlaceholder12#getPlaceholderId()} with
         * {@link Placeholder#HEADER} or {@link Placeholder#FOOTER}, to find out
         * what kind of placeholder this is.
         *
         * @return {@link RoundTripHFPlaceholder12} or {@code null} if not found
         *
         * @since POI 3.14-Beta2
         */
        public RoundTripHFPlaceholder12 getHFPlaceholderAtom()
        {
            // special case for files saved in Office 2007
            return GetClientDataRecord<RoundTripHFPlaceholder12>(RecordTypes.RoundTripHFPlaceholder12.typeID);
        }

        public bool isPlaceholder()
        {
            return
                ((getPlaceholderAtom() != null) ||
                //special case for files saved in Office 2007
                (getHFPlaceholderAtom() != null));
        }


        public IEnumerable<HSLFTextParagraph> iterator()
        {
            return _paragraphs;
        }

        /**
         * @since POI 5.2.0
         */
        public IEnumerable<HSLFTextParagraph> spliterator()
        {
            return _paragraphs;
        }

        //public Insets2D getInsets() {
        //    return new Insets2D(getTopInset(), getLeftInset(), getBottomInset(), getRightInset());
        //}

        //public void setInsets(Insets2D insets) {
        //    setTopInset(insets.top);
        //    setLeftInset(insets.left);
        //    setBottomInset(insets.bottom);
        //    setRightInset(insets.right);
        //}


        //public double getTextHeight() {
        //    return getTextHeight(null);
        //}


        //public double getTextHeight(Graphics2D graphics) {
        //    DrawFactory drawFact = DrawFactory.getInstance(graphics);
        //    DrawTextShape dts = drawFact.getDrawable(this);
        //    return dts.getTextHeight(graphics);
        //}


        public TextDirection getTextDirection()
        {
            // see 2.4.5 MSOTXFL
            AbstractEscherOptRecord opt = GetEscherOptRecord();
            EscherSimpleProperty prop = GetEscherProperty<EscherSimpleProperty>(opt, EscherProperties.TEXT__TEXTFLOW);
            int msotxfl = (prop == null) ? 0 : prop.PropertyValue;
            switch (msotxfl)
            {
                default:
                case 0: // msotxflHorzN
                case 4: // msotxflHorzA
                    return TextDirection.HORIZONTAL;
                case 1: // msotxflTtoBA
                case 3: // msotxflTtoBN
                case 5: // msotxflVertN
                    return TextDirection.VERTICAL;
                case 2: // msotxflBtoT
                    return TextDirection.VERTICAL_270;
                    // TextDirection.STACKED is not supported
            }
        }


        //public void setTextDirection(TextDirection orientation) {
        //    AbstractEscherOptRecord opt = GetEscherOptRecord();
        //    int msotxfl;
        //    if (orientation == null) {
        //        msotxfl = -1;
        //    } else {
        //        switch (orientation) {
        //            default:
        //            case TextDirection.STACKED:
        //                // not supported -> remove
        //                msotxfl = -1;
        //                break;
        //            case TextDirection.HORIZONTAL:
        //                msotxfl = 0;
        //                break;
        //            case TextDirection.VERTICAL:
        //                msotxfl = 1;
        //                break;
        //            case TextDirection.VERTICAL_270:
        //                // always interpreted as horizontal
        //                msotxfl = 2;
        //                break;
        //        }
        //    }
        //    setEscherProperty(opt, EscherProperties.TEXT__TEXTFLOW, msotxfl);
        //}


        //public Double getTextRotation() {
        //    // see 2.4.6 MSOCDIR
        //    AbstractEscherOptRecord opt = GetEscherOptRecord();
        //    EscherSimpleProperty prop = GetEscherProperty(opt, EscherProperties.TEXT__FONTROTATION);
        //    return (prop == null) ? null : (90 * prop.getPropertyValue());
        //}


        //public void setTextRotation(Double rotation) {
        //    AbstractEscherOptRecord opt = GetEscherOptRecord();
        //    if (rotation == null) {
        //        opt.RemoveEscherProperty(EscherProperties.TEXT__FONTROTATION);
        //    } else {
        //        int rot = (int)(Math.Round(rotation / 90) % 4L);
        //        setEscherProperty(EscherProperties.TEXT__FONTROTATION, rot);
        //    }
        //}

        /**
         * Returns the raw text content of the shape. This hasn't had any
         * changes applied to it, and so is probably unlikely to print
         * out nicely.
         */
        public String getRawText()
        {
            return HSLFTextParagraph.GetRawText(getTextParagraphs());
        }


        public String getText()
        {
            String rawText = getRawText();
            return HSLFTextParagraph.ToExternalString(rawText, getRunType());
        }


        //public HSLFTextRun appendText(String text, bool newParagraph) {
        //    // init paragraphs
        //    List<HSLFTextParagraph> paras = getTextParagraphs();
        //    HSLFTextRun htr = HSLFTextParagraph.AppendText(paras, text, newParagraph);
        //    setTextId(getRawText().GetHashCode());
        //    return htr;
        //}


        //public HSLFTextRun setText(String text) {
        //    // init paragraphs
        //    List<HSLFTextParagraph> paras = getTextParagraphs();
        //    HSLFTextRun htr = HSLFTextParagraph.setText(paras, text);
        //    setTextId(getRawText().hashCode());
        //    return htr;
        //}

        /**
         * Saves the modified paragraphs/textrun to the records.
         * Also updates the styles to the correct text length.
         */
        //public void storeText() {
        //    List<HSLFTextParagraph> paras = getTextParagraphs();
        //    HSLFTextParagraph.storeText(paras);
        //}

        /**
         * Returns the array of all hyperlinks in this text run
         *
         * @return the array of all hyperlinks in this text run or {@code null}
         *         if not found.
         */
        public List<HSLFHyperlink> getHyperlinks()
        {
            return HSLFHyperlink.Find(this);
        }


        //public void setTextPlaceholder(TextPlaceholder placeholder) {
        //    // TOOD: check for correct placeholder handling - see org.apache.poi.hslf.model.Placeholder
        //    Placeholder ph = null;
        //    int runType;
        //    switch (placeholder) {
        //        default:
        //        case BODY:
        //            runType = TextPlaceholder.BODY.nativeId;
        //            ph = Placeholder.BODY;
        //            break;
        //        case TITLE:
        //            runType = TITLE.nativeId;
        //            ph = Placeholder.TITLE;
        //            break;
        //        case CENTER_BODY:
        //            runType = TextPlaceholder.CENTER_BODY.nativeId;
        //            ph = Placeholder.BODY;
        //            break;
        //        case CENTER_TITLE:
        //            runType = CENTER_TITLE.nativeId;
        //            ph = Placeholder.TITLE;
        //            break;
        //        case HALF_BODY:
        //            runType = TextPlaceholder.HALF_BODY.nativeId;
        //            ph = Placeholder.BODY;
        //            break;
        //        case QUARTER_BODY:
        //            runType = TextPlaceholder.QUARTER_BODY.nativeId;
        //            ph = Placeholder.BODY;
        //            break;
        //        case NOTES:
        //            runType = TextPlaceholder.NOTES.nativeId;
        //            break;
        //        case OTHER:
        //            runType = TextPlaceholder.OTHER.nativeId;
        //            break;
        //    }
        //    setRunType(runType);
        //    if (ph != null) {
        //        setPlaceholder(ph);
        //    }
        //}


        public TextPlaceholder getTextPlaceholder()
        {
            return TextPlaceholder.fromNativeId(getRunType());
        }


        /**
         * Get alternative representation of text shape stored as metro blob escher property.
         * The returned shape is the first shape in stored group shape of the metro blob
         *
         * @return null, if there's no alternative representation, otherwise the text shape
         */
        //public <
        //    S : Shape<S,P>,
        //    P : TextParagraph<S,P,? : TextRun>
        //> Shape<S,P> getMetroShape() {
        //    return new HSLFMetroShape<S,P>(this).getShape();
        //}
    }
}