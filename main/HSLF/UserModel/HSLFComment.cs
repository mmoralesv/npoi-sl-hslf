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
using NPOI.HSLF.Record;
using NPOI.SL.UserModel;
using NPOI.Util;
using System;
using System.Collections.Generic;

namespace NPOI.HSLF.UserModel
{

    public class HSLFComment implements Comment {
    private Comment2000 _comment2000;

    public HSLFComment(Comment2000 comment2000) {
        _comment2000 = comment2000;
    }

    protected Comment2000 getComment2000() {
        return _comment2000;
    }

    /**
     * Get the Author of this comment
     */
    @Override
    public String getAuthor() {
        return _comment2000.getAuthor();
    }

    /**
     * Set the Author of this comment
     */
    @Override
    public void setAuthor(String author) {
        _comment2000.setAuthor(author);
    }

    /**
     * Get the Author's Initials of this comment
     */
    @Override
    public String getAuthorInitials() {
        return _comment2000.getAuthorInitials();
    }

    /**
     * Set the Author's Initials of this comment
     */
    @Override
    public void setAuthorInitials(String initials) {
        _comment2000.setAuthorInitials(initials);
    }

    /**
     * Get the text of this comment
     */
    @Override
    public String getText() {
        return _comment2000.getText();
    }

    /**
     * Set the text of this comment
     */
    @Override
    public void setText(String text) {
        _comment2000.setText(text);
    }

    @Override
    public Date getDate() {
        return _comment2000.getComment2000Atom().getDate();
    }

    @Override
    public void setDate(Date date) {
        _comment2000.getComment2000Atom().setDate(date);
    }

    @Override
    public Point2D getOffset() {
        double x = Units.masterToPoints(_comment2000.getComment2000Atom().getXOffset());
        double y = Units.masterToPoints(_comment2000.getComment2000Atom().getYOffset());
        return new Point2D.Double(x, y);
    }

    @Override
    public void setOffset(Point2D offset) {
        int x = Units.pointsToMaster(offset.getX());
        int y = Units.pointsToMaster(offset.getY());
        _comment2000.getComment2000Atom().setXOffset(x);
        _comment2000.getComment2000Atom().setYOffset(y);
    }
}
