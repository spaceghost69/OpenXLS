/*
 * --------- BEGIN COPYRIGHT NOTICE ---------
 * Copyright 2002-2012 Extentech Inc.
 * Copyright 2013 Infoteria America Corp.
 *
 * This file is part of OpenXLS.
 *
 * OpenXLS is free software: you can redistribute it and/or modify
 * it under the terms of the GNU Lesser General Public License as
 * published by the Free Software Foundation, either version 3 of
 * the License, or (at your option) any later version.
 *
 * OpenXLS is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 * GNU Lesser General Public License for more details.
 *
 * You should have received a copy of the GNU Lesser General Public
 * License along with OpenXLS.  If not, see
 * <http://www.gnu.org/licenses/>.
 * ---------- END COPYRIGHT NOTICE ----------
 */
package com.valkyrlabs.formats.LEO;

import com.valkyrlabs.toolkit.Logger;

import java.nio.ByteBuffer;
import java.util.ArrayList;

/**
 * The Root Storage == The Root Directory is the 1st Directory in the Directory Stream
 * This Directory is the root for all objects; it also stores the size and starting sector of the miniStream
 * <p>
 * The Root Directory entry behaves as both a stream and a storage object.  It's name MUST = "Root Entry"
 */
public class RootStorage extends com.valkyrlabs.formats.LEO.Storage {
    /**
     * serialVersionUID
     */
    private static final long serialVersionUID = -6568586717509723981L;
    int DEBUGLEVEL = 0;


    RootStorage(ByteBuffer b) {
        super(b);
    }

    public byte[] getBytes() {
        if (DEBUGLEVEL > 5) Logger.logInfo("Getting Root Storage Bytes....");
        return super.getBytes();
    }

    /**
     * set the underlying byte array for this
     * Storage
     * <p>
     * This appears to be a special case for RootStorage, as it seems to convert from a
     * smallblock based storage to a bigblock based storage.  works.... I guess?
     */
    public void setBytes(byte[] b) {
        Block[] bs = BlockFactory.getBlocksFromByteArray(b, Block.BIG);
        if (super.myblocks != null) {
            myblocks.clear();
            lastblock = null;
        } else {
            myblocks = new ArrayList(bs.length);
        }
        for (int d = 0; d < bs.length; d++)
            this.addBlock(bs[d]);
    }
}