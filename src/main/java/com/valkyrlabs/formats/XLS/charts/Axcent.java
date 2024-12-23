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
package com.valkyrlabs.formats.XLS.charts;

import com.valkyrlabs.formats.XLS.XLSRecord;
import com.valkyrlabs.toolkit.ByteTools;

/**
 * <b>Axcent: Axis Options(0x1062)</b>
 * <p>
 * 4		catMin		2		minimum date on axis.
 * If fAutoMin is set to 1, MUST be ignored.
 * If fDateAxis is set to 0, MUST be ignored.
 * 6		catMax		2		maximum date on axis.
 * fAutoMax is set to 1, MUST be ignored.
 * If fDateAxis is set to 0, MUST be ignored.
 * 8		catMajor	2		value of major unit
 * MUST be greater than or equal to catMinor when duMajor is equal to duMinor.
 * If fAutoMajor is set to 1, MUST be ignored.
 * If fDateAxis is set to 0, MUST be ignored.
 * 10		duMajor		2		Date Enumeration specifies unit of time for use of catMajor
 * If fDateAxis is set to 0, MUST be ignored.
 * 12		catMinor	2		value of minor unit
 * 14		duMinor		2		time units of minor unit
 * If fDateAxis is set to 0, MUST be ignored.
 * 16		duBase		2		smallest unit of time used by the axis.
 * If fAutoBase is set to 1, this field MUST be ignored.
 * If fDateAxis is set to 0, MUST be ignored.
 * 18		catCrossDate 2		crossing point of value axis (date)
 * 20		grbit		2
 * <p>
 * 0	0x1		fAutoMin	1= use default
 * 1	0x2		fAutoMax	""
 * 2	0x4		fAutoMajor	""
 * 3	0x8		fAutoMinor	""
 * 4	0x10	fdateAxis	1= this is a date axis
 * 5	0x20	fAutoBase	1= use default base
 * 6	0x40	fAutoCross	""
 * 7	0x80	fAutoDate	1= use default date settings for axis
 */
public class Axcent extends GenericChartObject implements ChartObject {
    /**
     * serialVersionUID
     */
    private static final long serialVersionUID = -660100252646337769L;
    private final byte[] PROTOTYPE_BYTES = new byte[]{0, 0, 0, 0, 1, 0, 0, 0, 1, 0, 0, 0, 0, 0, 0, 0, -17, 0};
    short catMin, catMax, catMajor, duMajor, catMinor, duMinor, duBase, catCrossDate, grbit;

    // 20070723 KSC: Need to create new records
    public static XLSRecord getPrototype() {
        Axcent a = new Axcent();
        a.setOpcode(AXCENT);
        a.setData(a.PROTOTYPE_BYTES);
        a.init();    // important when we parse these options ...
        return a;
    }

    public void init() {
        super.init();
        // 20071223 KSC: Start parsing of options
        catMin = ByteTools.readShort(this.getByteAt(0), this.getByteAt(1));
        catMax = ByteTools.readShort(this.getByteAt(2), this.getByteAt(3));
        catMajor = ByteTools.readShort(this.getByteAt(4), this.getByteAt(5));
        duMajor = ByteTools.readShort(this.getByteAt(6), this.getByteAt(7));
        catMinor = ByteTools.readShort(this.getByteAt(8), this.getByteAt(9));
        duMinor = ByteTools.readShort(this.getByteAt(10), this.getByteAt(11));
        duBase = ByteTools.readShort(this.getByteAt(12), this.getByteAt(13));
        catCrossDate = ByteTools.readShort(this.getByteAt(14), this.getByteAt(15));
        grbit = ByteTools.readShort(this.getByteAt(16), this.getByteAt(17));
    }

    // 20071223 KSC: add access methods
    public short getCatMin() {
        return catMin;
    }

    public short getCatMax() {
        return catMax;
    }

    public short getCatMajor() {
        return catMajor;
    }

    public short getDuMajor() {
        return duMajor;
    }

    public short getCatMinor() {
        return catMinor;
    }

    public short getDuMinor() {
        return duMinor;
    }

    public short getDuBase() {
        return duBase;
    }

    public boolean isDefaultMin() {
        return ((grbit & 0x1) == 0x1);
    }

    public boolean isDefaultMax() {
        return ((grbit & 0x2) == 0x2);
    }

    public boolean isDefaultMajorUnits() {
        return ((grbit & 0x4) == 0x4);
    }

    public boolean isDefaultMinorUnits() {
        return ((grbit & 0x8) == 0x8);
    }
}
