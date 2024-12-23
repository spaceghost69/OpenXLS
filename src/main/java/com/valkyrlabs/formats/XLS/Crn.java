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

import java.nio.charset.StandardCharsets;
import java.util.ArrayList;

/**
 * <b>CRN (005Ah)</b><br>
 * <p>
 * This record stores the contents of an external cell or cell range. An external cell range has one row only. If a cell range
 * spans over more than one row, several CRN records will be created.
 *
 * <p><pre>
 * offset  size    contents
 * ---
 * 0 		1 		Index to last column inside of the referenced sheet (lc)
 * 1 		1 		Index to first column inside of the referenced sheet (fc)
 * 2 		2 		Index to row inside of the referenced sheet
 * 4 		var.	List of lc-fc+1 cached values
 * </p></pre>
 */

public class Crn extends XLSRecord {
    /**
     * serialVersionUID
     */
    private static final long serialVersionUID = 3162130963170092322L;
    private final ArrayList cachedValues = new ArrayList();
    private byte lc, fc;
    private int rowIndex;

    public void init() {
        super.init();
        lc = this.getByteAt(0);
        fc = this.getByteAt(1);
        rowIndex = ByteTools.readShort(this.getByteAt(2), this.getByteAt(3));
        int pos = 4;
        for (int i = 0; i < lc - fc + 1; i++) {
            try {
                int type = this.getByteAt(pos++);
                switch (type) {
                    case 0:    // empty
                        pos += 8;
                        break;
                    case 1:    // numeric
                        cachedValues.add(new Float(ByteTools.eightBytetoLEDouble(this.getBytesAt(pos, 8))));
                        pos += 8;
                        break;
                    case 2: // string
                        short ln = ByteTools.readShort(this.getByteAt(pos), this.getByteAt(pos + 1));
                        byte encoding = this.getByteAt(pos + 2);
                        pos += 3;
                        if (encoding == 0) {
                            cachedValues.add(new String(this.getBytesAt(pos, ln)));
                            pos += ln;
                        } else {// unicode
                            cachedValues.add(new String(this.getBytesAt(pos, ln * 2), StandardCharsets.UTF_16LE));
                            pos += ln * 2;
                        }
                        break;
                    case 4: // boolean
                        cachedValues.add(Boolean.valueOf(this.getByteAt(pos + 1) == 1));
                        pos += 8;
                        break;
                    case 16: // error
                        cachedValues.add("Error Code: " + this.getByteAt(pos + 1));
                        pos += 8;
                        break;
                }
            } catch (Exception e) {

            }
        }
    }

    public String toString() {
        String ret = "CRN: lc=" + lc + " fc=" + fc + " rowIndex=" + rowIndex;
        for (int i = 0; i < cachedValues.size(); i++) {
            ret += " (" + i + ")=" + cachedValues.get(i);
        }
        return ret;
    }
}
