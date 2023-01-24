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
using NPOI.HSLF.Record;
using NPOI.SL.UserModel;
using NPOI.Util;
using System;
using System.Collections.Generic;

namespace NPOI.HSLF.UserModel
{

    /**
     * Create a <code>Shape</code> object depending on its type
     */
    public class HSLFShapeFactory {
    // For logging

    /**
     * Create a new shape from the data provided.
     */
    public static HSLFShape createShape(EscherContainerRecord spContainer, ShapeContainer<HSLFShape,HSLFTextParagraph> parent){
        if (spContainer.RecordId == EscherContainerRecord.SPGR_CONTAINER){
            return createShapeGroup(spContainer, parent);
        }
        return createSimpleShape(spContainer, parent);
    }

    public static HSLFGroupShape createShapeGroup(EscherContainerRecord spContainer, ShapeContainer<HSLFShape,HSLFTextParagraph> parent){
        bool isTable = false;
        EscherRecord child = spContainer.GetChild(0);
        if (!(child is EscherContainerRecord)) {
            throw new RecordFormatException("Did not have a EscherContainerRecord: " + child);
        }
        EscherContainerRecord ecr = (EscherContainerRecord) child;
        EscherRecord opt = HSLFShape.GetEscherChild(ecr, EscherProperties.USER_DEFINED) as EscherRecord;

        if (opt != null) {
            EscherPropertyFactory f = new EscherPropertyFactory();
            List<EscherProperty> props = f.CreateProperties( opt.Serialize(), 8, opt.Instance );
            foreach (EscherProperty ep in props) {
                if (ep.PropertyNumber == EscherProperties.GROUPSHAPE__TABLEPROPERTIES
                    && ep is EscherSimpleProperty
                    && (((EscherSimpleProperty)ep).PropertyValue & 1) == 1) {
                    isTable = true;
                    break;
                }
            }
        }

        HSLFGroupShape group;
        if (isTable) {
            group = new HSLFTable(spContainer, parent);

        } else {
            group = new HSLFGroupShape(spContainer, parent);
        }

        return group;
     }

    public static HSLFShape createSimpleShape(EscherContainerRecord spContainer, ShapeContainer<HSLFShape,HSLFTextParagraph> parent){
        EscherSpRecord spRecord = spContainer.GetChildById(EscherSpRecord.RECORD_ID) as EscherSpRecord;
        if (spRecord == null) {
            throw new RecordFormatException("Could not read EscherSpRecord as child of " + spContainer.RecordName);
        }

        HSLFShape shape;
        ShapeType type = ShapeType.forId(spRecord.getShapeType(), false);
        switch (type){
            case TEXT_BOX:
                shape = new HSLFTextBox(spContainer, parent);
                break;
            case HOST_CONTROL:
            case FRAME:
                shape = createFrame(spContainer, parent);
                break;
            case LINE:
                shape = new HSLFLine(spContainer, parent);
                break;
            case NOT_PRIMITIVE:
                shape = createNonPrimitive(spContainer, parent);
                break;
            default:
                if (parent is HSLFTable) {
                    EscherTextboxRecord etr = spContainer.getChildById(EscherTextboxRecord.RECORD_ID);
                    if (etr == null) {
                        LOG.atWarn().log("invalid ppt - add EscherTextboxRecord to cell");
                        etr = new EscherTextboxRecord();
                        etr.setRecordId(EscherTextboxRecord.RECORD_ID);
                        etr.setOptions((short)15);
                        spContainer.addChildRecord(etr);
                    }
                    shape = new HSLFTableCell(spContainer, (HSLFTable)parent);
                } else {
                    shape = new HSLFAutoShape(spContainer, parent);
                }
                break;
        }
        return shape;
    }

    private static HSLFShape createFrame(EscherContainerRecord spContainer, ShapeContainer<HSLFShape,HSLFTextParagraph> parent) {
        InteractiveInfo info = getClientDataRecord(spContainer, RecordTypes.InteractiveInfo.typeID);
        if(info != null && info.getInteractiveInfoAtom() != null){
            switch(info.getInteractiveInfoAtom().getAction()){
                case InteractiveInfoAtom.ACTION_OLE:
                    return new HSLFObjectShape(spContainer, parent);
                case InteractiveInfoAtom.ACTION_MEDIA:
                    return new MovieShape(spContainer, parent);
                default:
                    break;
            }
        }

        ExObjRefAtom oes = getClientDataRecord(spContainer, RecordTypes.ExObjRefAtom.typeID);
        return (oes != null)
            ? new HSLFObjectShape(spContainer, parent)
            : new HSLFPictureShape(spContainer, parent);
    }

    private static HSLFShape createNonPrimitive(EscherContainerRecord spContainer, ShapeContainer<HSLFShape,HSLFTextParagraph> parent) {
        AbstractEscherOptRecord opt = HSLFShape.getEscherChild(spContainer, EscherOptRecord.RECORD_ID);
        EscherProperty prop = HSLFShape.getEscherProperty(opt, EscherPropertyTypes.GEOMETRY__VERTICES);
        if(prop != null) {
            return new HSLFFreeformShape(spContainer, parent);
        }

        //LOG.atInfo().log("Creating AutoShape for a NotPrimitive shape");
        return new HSLFAutoShape(spContainer, parent);
    }

    protected static Record.Record getClientDataRecord(EscherContainerRecord spContainer, int recordType) {
        HSLFEscherClientDataRecord cldata = spContainer.GetChildById(EscherClientDataRecord.RECORD_ID);
        if (cldata != null) {
                foreach (Record.Record r in cldata.getHSLFChildRecords()) {
            if (r.getRecordType() == recordType) {
                return (T)r;
            }
        }
        return null;
    }
}
