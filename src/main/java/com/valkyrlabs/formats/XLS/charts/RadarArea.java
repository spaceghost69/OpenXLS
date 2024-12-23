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
 * <b>RadarArea: Chart Group Is a Radar Area Chart Group(0x1040)</b>
 * (i.e. a filled radar chart)
 * <p>
 * A - fRdrAxLab (1 bit): A bit that specifies whether category (3) labels are displayed.
 * <p>
 * B - fHasShadow (1 bit): A bit that specifies whether the data points in the chart group have shadows.
 * <p>
 * reserved (14 bits): MUST be zero, and MUST be ignored.
 * <p>
 * unused (2 bytes): Undefined and MUST be ignored.
 */
public class RadarArea extends GenericChartObject implements ChartObject {
    /**
     * serialVersionUID
     */
    private static final long serialVersionUID = 5731720802332312350L;
    private final byte[] PROTOTYPE_BYTES = new byte[]{1, 0};
    private short grbit = 0;
    private boolean fRdrAxLab = false;

    public static XLSRecord getPrototype() {
        RadarArea r = new RadarArea();
        r.setOpcode(RADARAREA);
        r.setData(r.PROTOTYPE_BYTES);
        r.init();
        return r;
    }

    public void init() {
        super.init();
        chartType = ChartConstants.RADARAREACHART;
        grbit = ByteTools.readShort(this.getByteAt(0), this.getByteAt(1));
        fRdrAxLab = (grbit & 0x1) == 0x1;
    }

    private void updateRecord() {
        byte[] b = ByteTools.shortToLEBytes(grbit);
        this.getData()[0] = b[0];
        this.getData()[1] = b[1];
    }

    /**
     * @return String XML representation of this chart-type's options
     */
    public String getOptionsXML() {
        StringBuffer sb = new StringBuffer();
        if (fRdrAxLab)
            sb.append(" AxisLabels=\"true\"");
        return sb.toString();
    }

    /**
     * Handle setting options from XML in a generic manner
     */
    public boolean setChartOption(String op, String val) {
        boolean bHandled = false;
        if (op.equalsIgnoreCase("AxisLabels")) {
            fRdrAxLab = val.equals("true");
            grbit = ByteTools.updateGrBit(grbit, fRdrAxLab, 0);
            bHandled = true;
        }
        if (bHandled)
            updateRecord();
        return bHandled;
    }
}
