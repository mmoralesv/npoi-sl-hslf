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
using NPOI.HSLF.Record;
using NPOI.SL.UserModel;
using NPOI.SS.Formula.Functions;
using NPOI.Util;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;

namespace NPOI.HSLF.UserModel
{
	/**
 * Represents a Shape which is the elemental object that composes a drawing.
 *  This class is a wrapper around EscherSpContainer which holds all information
 *  about a shape in PowerPoint document.
 *  <p>
 *  When you add a shape, you usually specify the dimensions of the shape and the position
 *  of the upper'left corner of the bounding box for the shape relative to the upper'left
 *  corner of the page, worksheet, or slide. Distances in the drawing layer are measured
 *  in points (72 points = 1 inch).
 */
	public class HSLFShape: Shape<HSLFShape, HSLFTextParagraph>
	{
		/**
     * Either EscherSpContainer or EscheSpgrContainer record
     * which holds information about this shape.
     */
		private EscherContainerRecord _escherContainer;

		/**
		 * Parent of this shape.
		 * {@code null} for the topmost shapes.
		 */
		private  ShapeContainer<HSLFShape, HSLFTextParagraph> _parent;

		/**
		 * The {@code Sheet} this shape belongs to
		 */
		private HSLFSheet _sheet;

		/**
		 * Fill
		 */
		private HSLFFill _fill;

		/**
		 * Create a Shape object. This constructor is used when an existing Shape is read from a PowerPoint document.
		 *
		 * @param escherRecord       {@code EscherSpContainer} container which holds information about this shape
		 * @param parent             the parent of this Shape
		 */
		protected HSLFShape(EscherContainerRecord escherRecord, ShapeContainer<HSLFShape, HSLFTextParagraph> parent)
		{
			_escherContainer = escherRecord;
			_parent = parent;
		}

		/**
		 * Create and assign the lower level escher record to this shape
		 */
		//protected EscherContainerRecord createSpContainer(bool isChild)
		//{
		//	if (_escherContainer == null)
		//	{
		//		_escherContainer = new EscherContainerRecord();
		//		_escherContainer.SetOptions((short)15);
		//	}
		//	return _escherContainer;
		//}

		/**
		 *  @return the parent of this shape
		 */
		//@Override
		public ShapeContainer<HSLFShape, HSLFTextParagraph> GetParent()
		{
			return _parent;
		}

		/**
		 * @return name of the shape.
		 */
		//@Override
		public String GetShapeName()
		{
			EscherComplexProperty ep = GetEscherOptRecord().GetEscherProperty(EscherProperties.GROUPSHAPE__SHAPENAME) as EscherComplexProperty;
			if (ep != null)
			{
				byte[] cd = ep.ComplexData;
				return StringUtil.GetFromUnicodeLE0Terminated(cd, 0, cd.Length / 2);
			}
			else
			{
				return GetShapeType().nativeName + " " + GetShapeId();
			}
		}

		public ShapeType GetShapeType()
		{
			EscherSpRecord spRecord = GetEscherChild(EscherSpRecord.RECORD_ID) as EscherSpRecord;
			return ShapeType.forId(spRecord.ShapeType, false);
		}

		public void SetShapeType(ShapeType type)
		{
			EscherSpRecord spRecord = GetEscherChild(EscherSpRecord.RECORD_ID) as EscherSpRecord;
			spRecord.ShapeType = (short)type.nativeId;
			spRecord.Version = 0x2;
		}

		/**
		 * Returns the anchor (the bounding box rectangle) of this shape.
		 * All coordinates are expressed in points (72 dpi).
		 *
		 * @return the anchor of this shape
		 */
		//@Override
		//public Rectangle2D GetAnchor()
		//{
		//	EscherSpRecord spRecord = GetEscherChild(EscherSpRecord.RECORD_ID);
		//	int flags = spRecord.GetFlags();
		//	int x1, y1, x2, y2;
		//	EscherChildAnchorRecord childRec = GetEscherChild(EscherChildAnchorRecord.RECORD_ID);
		//	bool useChildRec = ((flags & EscherSpRecord.FLAG_CHILD) != 0);
		//	if (useChildRec && childRec != null)
		//	{
		//		x1 = childRec.GetDx1();
		//		y1 = childRec.GetDy1();
		//		x2 = childRec.GetDx2();
		//		y2 = childRec.GetDy2();
		//	}
		//	else
		//	{
		//		if (useChildRec)
		//		{
		//			LOG.atWarn().log("EscherSpRecord.FLAG_CHILD is set but EscherChildAnchorRecord was not found");
		//		}
		//		EscherClientAnchorRecord clientRec = getEscherChild(EscherClientAnchorRecord.RECORD_ID);
		//		if (clientRec == null)
		//		{
		//			throw new RecordFormatException("Could not read record 'CLIENT_ANCHOR' with record-id: " + EscherClientAnchorRecord.RECORD_ID);
		//		}
		//		x1 = clientRec.getCol1();
		//		y1 = clientRec.getFlag();
		//		x2 = clientRec.getDx1();
		//		y2 = clientRec.getRow1();
		//	}

		//	// TODO: find out where this -1 value comes from at #57820 (link to ms docs?)

		//	return new Rectangle2D.Double(
		//		(x1 == -1 ? -1 : Units.masterToPoints(x1)),
		//		(y1 == -1 ? -1 : Units.masterToPoints(y1)),
		//		(x2 == -1 ? -1 : Units.masterToPoints(x2 - x1)),
		//		(y2 == -1 ? -1 : Units.masterToPoints(y2 - y1))
		//	);
		//}

		/**
		 * Sets the anchor (the bounding box rectangle) of this shape.
		 * All coordinates should be expressed in points (72 dpi).
		 *
		 * @param anchor new anchor
		 */
		//public void setAnchor(Rectangle2D anchor)
		//{
		//	int x = Units.pointsToMaster(anchor.getX());
		//	int y = Units.pointsToMaster(anchor.getY());
		//	int w = Units.pointsToMaster(anchor.getWidth() + anchor.getX());
		//	int h = Units.pointsToMaster(anchor.getHeight() + anchor.getY());
		//	EscherSpRecord spRecord = getEscherChild(EscherSpRecord.RECORD_ID);
		//	int flags = spRecord.getFlags();
		//	if ((flags & EscherSpRecord.FLAG_CHILD) != 0)
		//	{
		//		EscherChildAnchorRecord rec = getEscherChild(EscherChildAnchorRecord.RECORD_ID);
		//		rec.setDx1(x);
		//		rec.setDy1(y);
		//		rec.setDx2(w);
		//		rec.setDy2(h);
		//	}
		//	else
		//	{
		//		EscherClientAnchorRecord rec = getEscherChild(EscherClientAnchorRecord.RECORD_ID);
		//		rec.setCol1((short)x);
		//		rec.setFlag((short)y);
		//		rec.setDx1((short)w);
		//		rec.setRow1((short)h);
		//	}

		//}

		/**
		 * Moves the top left corner of the shape to the specified point.
		 *
		 * @param x the x coordinate of the top left corner of the shape
		 * @param y the y coordinate of the top left corner of the shape
		 */
		public  void MoveTo(double x, double y)
		{
			//// This convenience method should be implemented via setAnchor in subclasses
			//// see HSLFGroupShape.setAnchor() for a reference
			//Rectangle2D anchor = getAnchor();
			//anchor.setRect(x, y, anchor.getWidth(), anchor.getHeight());
			//setAnchor(anchor);
		}

		/**
		 * Helper method to return escher child by record ID
		 *
		 * @return escher record or {@code null} if not found.
		 */
		public static EscherRecord GetEscherChild(EscherContainerRecord owner, int recordId)
		{
			return owner.GetChildById((short)recordId);
		}

		public EscherRecord GetEscherChild(int recordId)
		{
			return _escherContainer.GetChildById((short)recordId);
		}

		/**
			 * Returns  escher property by type.
			 *
			 * @return escher property or {@code null} if not found.
			 */
		public static  T GetEscherProperty<T>(AbstractEscherOptRecord opt, int type)where T: EscherProperty
		{
			return (T)((opt == null) ? null : opt.Lookup(type));
		}


		/**
		 * Get the value of a simple escher property for this shape.
		 *
		 * @param propId    The id of the property. One of the constants defined in EscherOptRecord.
		 *
		 * @deprecated use {@link #getEscherProperty(EscherPropertyTypes, int)} instead
		 */
		//@Deprecated
		//@Removal(version = "5.0.0")
		public int GetEscherProperty(short propId, int defaultValue = 0)
		{
            AbstractEscherOptRecord opt = GetEscherOptRecord();
            EscherSimpleProperty prop = opt.GetEscherProperty(propId) as EscherSimpleProperty;
            return prop == null ? defaultValue : prop.PropertyValue;
        }

		/**
		 * @return  The shape container and its children that can represent this
		 *          shape.
		 */
		public EscherContainerRecord GetSpContainer()
		{
			return _escherContainer;
		}

		/**
		 * Event which fires when a shape is inserted in the sheet.
		 * In some cases we need to propagate changes to upper level containers.
		 * <br>
		 * Default implementation does nothing.
		 *
		 * @param sh - owning shape
		 */
		protected void AfterInsert(HSLFSheet sh)
		{
			throw new NotImplementedException();
			//if (_fill != null)
			//{
			//	_fill.afterInsert(sh);
			//}
		}

		/**
		 *  @return the {@code SlideShow} this shape belongs to
		 */
		//@Override
		public HSLFSheet GetSheet()
		{
			return _sheet;
		}

		/**
		 * Assign the {@code SlideShow} this shape belongs to
		 *
		 * @param sheet owner of this shape
		 */
		public void SetSheet(HSLFSheet sheet)
		{
			_sheet = sheet;
		}

		//@Override
		public int GetShapeId()
		{
			EscherSpRecord spRecord = (EscherSpRecord)GetEscherChild(EscherSpRecord.RECORD_ID);
			return spRecord == null ? 0 : spRecord.ShapeId;
		}

		/**
		 * Sets shape ID
		 *
		 * @param id of the shape
		 */
		public void SetShapeId(int id)
		{
			EscherSpRecord spRecord = (EscherSpRecord)GetEscherChild(EscherSpRecord.RECORD_ID);
			if (spRecord != null) spRecord.ShapeId = (id);
		}

		/**
		 * Fill properties of this shape
		 *
		 * @return fill properties of this shape
		 */
		public HSLFFill GetFill()
		{
			if (_fill == null)
			{
				_fill = new HSLFFill(this);
			}
			return _fill;
		}

		//@Override
		//public void Draw(Graphics2D graphics, Rectangle2D bounds)
		//{
		//	DrawFactory.getInstance(graphics).drawShape(graphics, this, bounds);
		//}

		public AbstractEscherOptRecord GetEscherOptRecord()
		{
			AbstractEscherOptRecord opt = GetEscherChild(EscherProperties.OPT) as AbstractEscherOptRecord;
			if (opt == null)
			{
				opt = GetEscherChild(EscherProperties.USER_DEFINED) as AbstractEscherOptRecord;
			}
			return opt;
		}

		public bool IsPlaceholder()
		{
			return false;
		}

		/**
		 *  Find a record in the underlying EscherClientDataRecord
		 *
		 * @param recordType type of the record to search
		 */
		//@SuppressWarnings("unchecked")

		public T GetClientDataRecord<T>(int recordType) where T : Record.Record
		{

			List <Record.Record > records = GetClientRecords();
			if (records != null) foreach (Record.Record r in records)
				{
					if (r.GetRecordType() == recordType)
					{
						return (T)r;
					}
				}
			return null;
		}

		/**
		 * Search for EscherClientDataRecord, if found, convert its contents into an array of HSLF records
		 *
		 * @return an array of HSLF records contained in the shape's EscherClientDataRecord or {@code null}
		 */
		protected List<Record.Record> GetClientRecords()
		{
			HSLFEscherClientDataRecord clientData = GetClientData(false);
			return (clientData == null) ? null : clientData.getHSLFChildRecords();
		}

		/**
		 * Create a new HSLF-specific EscherClientDataRecord
		 *
		 * @param create if true, create the missing record
		 * @return the client record or null if it was missing and create wasn't activated
		 */
		protected HSLFEscherClientDataRecord GetClientData(bool create)
		{
			HSLFEscherClientDataRecord clientData = GetEscherChild(EscherClientDataRecord.RECORD_ID) as HSLFEscherClientDataRecord;
			if (clientData == null && create)
			{
				clientData = new HSLFEscherClientDataRecord();
				clientData.Options = 15;
				clientData.RecordId = EscherClientDataRecord.RECORD_ID;
				GetSpContainer().AddChildBefore(clientData, EscherTextboxRecord.RECORD_ID);
			}
			return clientData;
		}
	}
}
