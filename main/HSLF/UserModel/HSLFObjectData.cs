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

    /**
     * A class that represents object data embedded in a slide show.
     */
    public class HSLFObjectData implements ObjectData, GenericRecord {
    /**
     * The record that contains the object data.
     */
    private ExOleObjStg storage;

    /**
     * Creates the object data wrapping the record that contains the object data.
     *
     * @param storage the record that contains the object data.
     */
    public HSLFObjectData(ExOleObjStg storage) {
        this.storage = storage;
    }

    @Override
    public InputStream getInputStream() {
        return storage.getData();
    }

    @Override
    public OutputStream getOutputStream() {
        // can't use UnsynchronizedByteArrayOutputStream here, because it's final
        return new ByteArrayOutputStream() {
            @Override
            public void close() throws IOException {
                setData(getBytes());
            }
        };
    }

    /**
     * Sets the embedded data.
     *
     * @param data the embedded data.
     */
     public void setData(byte[] data) throws IOException {
        storage.setData(data);
    }

    /**
     * Return the record that contains the object data.
     *
     * @return the record that contains the object data.
     */
    public ExOleObjStg getExOleObjStg() {
        return storage;
    }


    @Override
    public String getOLE2ClassName() {
        return null;
    }

    @Override
    public String getFileName() {
        return null;
    }

    @Override
    public Map<String, Supplier<?>> getGenericProperties() {
        return null;
    }

    @Override
    public List<? : GenericRecord> getGenericChildren() {
        return Collections.singletonList(getExOleObjStg());
    }
}
