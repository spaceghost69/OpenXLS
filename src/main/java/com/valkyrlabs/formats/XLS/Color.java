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
package com.valkyrlabs.formats.XLS;

import com.valkyrlabs.toolkit.ByteTools;
import com.valkyrlabs.toolkit.Logger;

import java.io.Serializable;


/**
 * An RGB Color Value
 */

public final class Color implements Serializable {

    /**
     * serialVersionUID
     */
    private static final long serialVersionUID = 5181253361814526629L;
    byte[] data = null;
    int myval = -1;

    public Color(byte[] d) {
        this.data = d;
        myval = ByteTools.readInt(d[0], d[1], d[2], d[3]);
        if (false) Logger.logInfo("New Color: " + Integer.toHexString(myval));
    }

    // methods from BiffRec

    public boolean isDirty(Object obj) {
        return false; //immutable for now...
    }

    public byte[] read() {
        return this.data;
    }
}