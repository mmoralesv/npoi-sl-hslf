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
using NPOI.HSLF.Exceptions;
using NPOI.HSLF.Model;
using NPOI.HSLF.Record;
using NPOI.SL.UserModel;
using NPOI.SS.Formula.Functions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using static NPOI.HSLF.Record.SlideAtomLayout;
using static NPOI.HSLF.Record.SlideListWithText;

namespace NPOI.HSLF.UserModel
{
	public class HSLFSlide: HSLFSheet, Slide<HSLFShape, HSLFTextParagraph>
	{
		private int _slideNo;
		private SlideAtomsSet _atomSet;
		private  List<List<HSLFTextParagraph>> _paragraphs = new List<List<HSLFTextParagraph>>();
		private HSLFNotes _notes; // usermodel needs to set this

		public MasterSheet<HSLFShape, HSLFTextParagraph> MasterSheet { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
		public MasterSheet<HSLFShape, HSLFTextParagraph> SlideLayout { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }

		/**
		 * Constructs a Slide from the Slide record, and the SlideAtomsSet
		 *  containing the text.
		 * Initializes TextRuns, to provide easier access to the text
		 *
		 * @param slide the Slide record we're based on
		 * @param notes the Notes sheet attached to us
		 * @param atomSet the SlideAtomsSet to get the text from
		 */
		public HSLFSlide(Record.Slide slide, HSLFNotes notes, SlideAtomsSet atomSet, int slideIdentifier, int slideNumber)
			:base(slide, slideIdentifier)
		{
			_notes = notes;
			_atomSet = atomSet;
			_slideNo = slideNumber;

			// For the text coming in from the SlideAtomsSet:
			// Build up TextRuns from pairs of TextHeaderAtom and
			//  one of TextBytesAtom or TextCharsAtom
			if (_atomSet != null && _atomSet.getSlideRecords().Length > 0)
			{
				// Grab text from SlideListWithTexts entries
				_paragraphs.AddRange(HSLFTextParagraph.FindTextParagraphs(_atomSet.getSlideRecords()));
				if (_paragraphs.Count ==0)
				{
					throw new HSLFException("No text records found for slide");
				}
			}

			// Grab text from slide's PPDrawing
			foreach (List<HSLFTextParagraph> l in HSLFTextParagraph.FindTextParagraphs(GetPPDrawing(), this))
			{
				if (!_paragraphs.Contains(l))
				{
					_paragraphs.Add(l);
				}
			}
		}

		/**
		* Create a new Slide instance
		* @param sheetNumber The internal number of the sheet, as used by PersistPtrHolder
		* @param slideNumber The user facing number of the sheet
		*/
		public HSLFSlide(int sheetNumber, int sheetRefId, int slideNumber)
			:base(new Record.Slide(), sheetNumber)
		{
			_slideNo = slideNumber;
			GetSheetContainer().SetSheetId(sheetRefId);
		}

		/**
		 * Returns the Notes Sheet for this slide, or null if there isn't one
		 */
		//@Override
	public HSLFNotes GetNotes()
		{
			return _notes;
		}

		/**
		 * Sets the Notes that are associated with this. Updates the
		 *  references in the records to point to the new ID
		 */
		//@Override
	public void SetNotes(Notes<HSLFShape, HSLFTextParagraph> notes)
		{
			if (notes != null && !(notes is HSLFNotes)) {
				throw new ArgumentException("notes needs to be of type HSLFNotes");
			}
			_notes = (HSLFNotes)notes;

			// Update the Slide Atom's ID of where to point to
			SlideAtom sa = getSlideRecord().getSlideAtom();

			if (_notes == null)
			{
				// Set to 0
				sa.setNotesID(0);
			}
			else
			{
				// Set to the value from the notes' sheet id
				sa.setNotesID(_notes._getSheetNumber());
			}
		}

		/**
		* Changes the Slide's (external facing) page number.
		* @see org.apache.poi.hslf.usermodel.HSLFSlideShow#reorderSlide(int, int)
		*/
		public void setSlideNumber(int newSlideNumber)
		{
			_slideNo = newSlideNumber;
		}

		/**
		 * Called by SlideShow ater a new slide is created.
		 * <p>
		 * For Slide we need to do the following:
		 * <ul>
		 *  <li> set id of the drawing group.</li>
		 *  <li> set shapeId for the container descriptor and background</li>
		 * </ul>
		 */
		//@Override
	//public void onCreate()
	//	{
	//		//initialize drawing group id
	//		EscherDggRecord dgg = GetSlideShow().getDocumentRecord().GetPPDrawingGroup().getEscherDggRecord();
	//		EscherContainerRecord dgContainer = GetSheetContainer().GetPPDrawing().getDgContainer();
	//		EscherDgRecord dg = (EscherDgRecord)HSLFShape.GetEscherChild(dgContainer, EscherDgRecord.RECORD_ID);
	//		int dgId = dgg.getMaxDrawingGroupId() + 1;
	//		dg.setOptions((short)(dgId << 4));
	//		dgg.setDrawingsSaved(dgg.getDrawingsSaved() + 1);

	//		for (EscherContainerRecord c : dgContainer.getChildContainers())
	//		{
	//			EscherSpRecord spr = null;
	//			switch (EscherRecordTypes.forTypeID(c.getRecordId()))
	//			{
	//				case SPGR_CONTAINER:
	//					EscherContainerRecord dc = (EscherContainerRecord)c.getChild(0);
	//					spr = dc.getChildById(EscherSpRecord.RECORD_ID);
	//					break;
	//				case SP_CONTAINER:
	//					spr = c.getChildById(EscherSpRecord.RECORD_ID);
	//					break;
	//				default:
	//					break;
	//			}
	//			if (spr != null)
	//			{
	//				spr.setShapeId(allocateShapeId());
	//			}
	//		}

	//		//PPT doen't increment the number of saved shapes for group descriptor and background
	//		dg.setNumShapes(1);
	//	}

		/**
		 * Create a {@code TextBox} object that represents the slide's title.
		 *
		 * @return {@code TextBox} object that represents the slide's title.
		 */
		//public HSLFTextBox addTitle()
		//{
		//	HSLFPlaceholder pl = new HSLFPlaceholder();
		//	pl.setShapeType(ShapeType.RECT);
		//	pl.setPlaceholder(Placeholder.TITLE);
		//	pl.setRunType(TextPlaceholder.TITLE.nativeId);
		//	pl.setText("Click to edit title");
		//	pl.setAnchor(new java.awt.Rectangle(54, 48, 612, 90));
		//	addShape(pl);
		//	return pl;
		//}


		// Complex Accesser methods follow

		/**
		 * <p>
		 * The title is a run of text of type {@code TextHeaderAtom.CENTER_TITLE_TYPE} or
		 * {@code TextHeaderAtom.TITLE_TYPE}
		 * </p>
		 *
		 * @see TextHeaderAtom
		 */
		//@Override
	public String GetTitle()
		{
			foreach (List<HSLFTextParagraph> tp in getTextParagraphs())
			{
				if (tp.Count==0)
				{
					continue;
				}
				int type = tp.ElementAt(0).GetRunType();
				if (TextPlaceholder.isTitle(type))
				{
					String str = HSLFTextParagraph.GetRawText(tp);
					return HSLFTextParagraph.ToExternalString(str, type);
				}
			}
			return null;
		}

		//@Override
	public String GetSlideName()
		{
			 CString name = (CString)getSlideRecord().FindFirstOfType(RecordTypes.CString.typeID);
			return name != null ? name.GetText() : "Slide" + GetSlideNumber();
		}


		/**
		 * Returns an array of all the TextRuns found
		 */
		//@Override
	public List<List<HSLFTextParagraph>> getTextParagraphs() { return _paragraphs; }

		/**
		 * Returns the (public facing) page number of this slide
		 */
		//@Override
	public int GetSlideNumber() { return _slideNo; }

		/**
		 * Returns the underlying slide record
		 */
		public Record.Slide getSlideRecord()
		{
			return (Record.Slide)GetSheetContainer();
		}

		/**
		 * @return set of records inside {@code SlideListWithtext} container
		 *  which hold text data for this slide (typically for placeholders).
		 */
		public SlideAtomsSet getSlideAtomsSet() { return _atomSet; }

		/**
		 * Returns master sheet associated with this slide.
		 * It can be either SlideMaster or TitleMaster objects.
		 *
		 * @return the master sheet associated with this slide.
		 */
		//@Override
	public HSLFMasterSheet getMasterSheet()
		{
			int masterId = getSlideRecord().getSlideAtom().getMasterID();
			foreach (HSLFSlideMaster sm in GetSlideShow().getSlideMasters())
			{
				if (masterId == sm._getSheetNumber())
				{
					return sm;
				}
			}
			foreach (HSLFTitleMaster tm in GetSlideShow().getTitleMasters())
			{
				if (masterId == tm._getSheetNumber())
				{
					return tm;
				}
			}
			return null;
		}

		/**
		 * Change Master of this slide.
		 */
		public void setMasterSheet(HSLFMasterSheet master)
		{
			SlideAtom sa = getSlideRecord().getSlideAtom();
			int sheetNo = master._getSheetNumber();
			sa.setMasterID(sheetNo);
		}

		/**
		 * Sets whether this slide follows master background
		 *
		 * @param flag  {@code true} if the slide follows master,
		 * {@code false} otherwise
		 */
		//@Override
	public void setFollowMasterBackground(bool flag)
		{
			SlideAtom sa = getSlideRecord().getSlideAtom();
			sa.setFollowMasterBackground(flag);
		}

		/**
		 * Whether this slide follows master sheet background
		 *
		 * @return {@code true} if the slide follows master background,
		 * {@code false} otherwise
		 */
		//@Override
	public bool getFollowMasterBackground()
		{
			SlideAtom sa = getSlideRecord().getSlideAtom();
			return sa.getFollowMasterBackground();
		}

		/**
		 * Sets whether this slide draws master sheet objects
		 *
		 * @param flag  {@code true} if the slide draws master sheet objects,
		 * {@code false} otherwise
		 */
		//@Override
	public void setFollowMasterObjects(bool flag)
		{
			SlideAtom sa = getSlideRecord().getSlideAtom();
			sa.setFollowMasterObjects(flag);
		}

		/**
		 * Whether this slide follows master color scheme
		 *
		 * @return {@code true} if the slide follows master color scheme,
		 * {@code false} otherwise
		 */
		public bool getFollowMasterScheme()
		{
			SlideAtom sa = getSlideRecord().getSlideAtom();
			return sa.getFollowMasterScheme();
		}

		/**
		 * Sets whether this slide draws master color scheme
		 *
		 * @param flag  {@code true} if the slide draws master color scheme,
		 * {@code false} otherwise
		 */
		public void setFollowMasterScheme(bool flag)
		{
			SlideAtom sa = getSlideRecord().getSlideAtom();
			sa.setFollowMasterScheme(flag);
		}

		/**
		 * Whether this slide draws master sheet objects
		 *
		 * @return {@code true} if the slide draws master sheet objects,
		 * {@code false} otherwise
		 */
		//@Override
	public bool getFollowMasterObjects()
		{
			SlideAtom sa = getSlideRecord().getSlideAtom();
			return sa.getFollowMasterObjects();
		}

		/**
		 * Background for this slide.
		 */
		//@Override
	public HSLFBackground getBackground()
		{
			if (getFollowMasterBackground())
			{
				 HSLFMasterSheet ms = getMasterSheet();
				return (ms == null) ? null : ms.GetBackground();
			}
			return base.GetBackground();
		}

		/**
		 * Color scheme for this slide.
		 */
		//@Override
	public ColorSchemeAtom getColorScheme()
		{
			if (getFollowMasterScheme())
			{
				 HSLFMasterSheet ms = getMasterSheet();
				return (ms == null) ? null : ms.GetColorScheme();
			}
			return base.GetColorScheme();
		}

		private static RecordContainer selectContainer( RecordContainer root,  int index,  params RecordTypes[] path)
		{
			if (root == null || index >= path.Length)
			{
				return root;
			}
			 RecordContainer newRoot = (RecordContainer)root.FindFirstOfType(path[index].typeID);
			return selectContainer(newRoot, index + 1, path);
		}

		/**
		 * Get the comment(s) for this slide.
		 * Note - for now, only works on PPT 2000 and
		 *  PPT 2003 files. Doesn't work for PPT 97
		 *  ones, as they do their comments oddly.
		 */
		//@Override
	public List<HSLFComment> getComments()
		{
			List<HSLFComment> comments = new List<HSLFComment>();
			// If there are any, they're in
			//  ProgTags -> ProgBinaryTag -> BinaryTagData
			 RecordContainer binaryTags =
					selectContainer(GetSheetContainer(), 0,
							RecordTypes.ProgTags, RecordTypes.ProgBinaryTag, RecordTypes.BinaryTagData);

			if (binaryTags != null)
			{
				foreach ( Record.Record record in binaryTags.GetChildRecords())
				{
					if (record is Comment2000) {
					comments.Add(new HSLFComment((Comment2000)record));
				}
			}
		}

        return comments;
    }

	/**
     * Header / Footer settings for this slide.
     *
     * @return Header / Footer settings for this slide
     */
	//@Override
	public HeadersFooters getHeadersFooters()
	{
		return new HeadersFooters(this, HeadersFootersContainer.SlideHeadersFootersContainer);
	}

	//@Override
	protected void onAddTextShape(HSLFTextShape shape)
	{
		List<HSLFTextParagraph> newParas = shape.getTextParagraphs();
		_paragraphs.Add(newParas);
	}

	/** This will return an atom per TextBox, so if the page has two text boxes the method should return two atoms. */
	public StyleTextProp9Atom[] getNumberedListInfo()
	{
		return this.GetPPDrawing().getNumberedListInfo();
	}

	public EscherTextboxWrapper[] getTextboxWrappers()
	{
		return this.GetPPDrawing().GetTextboxWrappers();
	}

	//@Override
	public void setHidden(bool hidden)
	{
		Record.Slide cont = getSlideRecord();

		SSSlideInfoAtom slideInfo =
			(SSSlideInfoAtom)cont.FindFirstOfType(RecordTypes.SSSlideInfoAtom.typeID);
		if (slideInfo == null)
		{
			slideInfo = new SSSlideInfoAtom();
			cont.AddChildAfter(slideInfo, cont.FindFirstOfType(RecordTypes.SlideAtom.typeID));
		}

		slideInfo.setEffectTransitionFlagByBit(SSSlideInfoAtom.HIDDEN_BIT, hidden);
	}

	//@Override
	public bool isHidden()
	{
		SSSlideInfoAtom slideInfo =
			(SSSlideInfoAtom)getSlideRecord().FindFirstOfType(RecordTypes.SSSlideInfoAtom.typeID);
		return (slideInfo != null) && slideInfo.getEffectTransitionFlagByBit(SSSlideInfoAtom.HIDDEN_BIT);
	}

	//@Override
	//public void draw(Graphics2D graphics)
	//{
	//	DrawFactory drawFact = DrawFactory.getInstance(graphics);
	//	Drawable draw = drawFact.getDrawable(this);
	//	draw.draw(graphics);
	//}

	//@Override
	public bool getFollowMasterColourScheme()
	{
		return false;
	}

	//@Override
	public void setFollowMasterColourScheme(bool follow)
	{
	}

	//@Override
	public bool getFollowMasterGraphics()
	{
		return getFollowMasterObjects();
	}

	//TODO: Fix SlideAtomLayout
	//public bool getDisplayPlaceholder( Placeholder placeholder)
	//{
	//	 HeadersFooters hf = getHeadersFooters();
	//	 SlideLayoutType slt = getSlideRecord().getSlideAtom().getSSlideLayoutAtom().getGeometryType();
	//	 bool isTitle =
	//		(slt == SlideLayoutType.TITLE_SLIDE || slt == SlideLayoutType.TITLE_ONLY || slt == SlideLayoutType.MASTER_TITLE);
	//	switch (placeholder)
	//	{
	//		case DATETIME:
	//			return (hf.isDateTimeVisible() && (hf.isTodayDateVisible() || (hf.isUserDateVisible() && hf.getUserDateAtom() != null))) && !isTitle;
	//		case SLIDE_NUMBER:
	//			return hf.isSlideNumberVisible() && !isTitle;
	//		case HEADER:
	//			return hf.isHeaderVisible() && hf.getHeaderAtom() != null && !isTitle;
	//		case FOOTER:
	//			return hf.isFooterVisible() && hf.getFooterAtom() != null && !isTitle;
	//		default:
	//			return false;
	//	}
	//}

	//@Override
	//public bool getDisplayPlaceholder( SimpleShape<?,?> placeholderRef)
	//{
	//	Placeholder placeholder = placeholderRef.getPlaceholder();
	//	if (placeholder == null)
	//	{
	//		return false;
	//	}

	//	 HeadersFooters hf = getHeadersFooters();
	//	 SlideLayoutType slt = getSlideRecord().getSlideAtom().getSSlideLayoutAtom().getGeometryType();
	//	 bool isTitle =
	//		(slt == SlideLayoutType.TITLE_SLIDE || slt == SlideLayoutType.TITLE_ONLY || slt == SlideLayoutType.MASTER_TITLE);
	//	switch (placeholder)
	//	{
	//		case HEADER:
	//			return hf.isHeaderVisible() && hf.getHeaderAtom() != null && !isTitle;
	//		case FOOTER:
	//			return hf.isFooterVisible() && hf.getFooterAtom() != null && !isTitle;
	//		case DATETIME:
	//		case SLIDE_NUMBER:
	//		default:
	//			return false;
	//	}
	//}

	//@Override
	public HSLFMasterSheet getSlideLayout()
	{
		// TODO: find out how we can find the mastersheet base on the slide layout type, i.e.
		// getSlideRecord().getSlideAtom().getSSlideLayoutAtom().getGeometryType()
		return getMasterSheet();
	}

		List<Comment> Slide<HSLFShape, HSLFTextParagraph>.getComments()
		{
			throw new NotImplementedException();
		}

		Notes<HSLFShape, HSLFTextParagraph> Slide<HSLFShape, HSLFTextParagraph>.GetNotes()
		{
			throw new NotImplementedException();
		}

		public bool GetFollowMasterBackground()
		{
			throw new NotImplementedException();
		}

		public void SetFollowMasterBackground(bool follow)
		{
			throw new NotImplementedException();
		}

		public bool GetFollowMasterColourScheme()
		{
			throw new NotImplementedException();
		}

		public void SetFollowMasterColourScheme(bool follow)
		{
			throw new NotImplementedException();
		}

		public bool GetFollowMasterObjects()
		{
			throw new NotImplementedException();
		}

		public void SetFollowMasterObjects(bool follow)
		{
			throw new NotImplementedException();
		}

		public bool GetDisplayPlaceholder(SimpleShape<HSLFShape, HSLFTextParagraph> placeholderRefShape)
		{
			throw new NotImplementedException();
		}

		public void SetHidden(bool hidden)
		{
			throw new NotImplementedException();
		}

		public bool IsHidden()
		{
			throw new NotImplementedException();
		}
	}
}
