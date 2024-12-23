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

import com.valkyrlabs.formats.XLS.WorkBookFactory;
import com.valkyrlabs.toolkit.Logger;

import java.io.UnsupportedEncodingException;


/**
 * <b>SeriesText: Chart Legend/Category/Value Text Definition (100Dh)</b><br>
 * <p>
 * This record defines the SeriesText data of a chart.
 * <p>
 * sdtX and sdtY fields determine data type (numeric and text)
 * <p>
 * cValx and cValy fields determine number of cells in series
 * <p>
 * Offset           Name    Size    Contents
 * --
 * 4               id    	2       Text identifier (should be zero)
 * 8               cch     2       length of String text
 * 10              rgch    2       String text
 *
 * </pre>
 *
 * @see Chart
 */

public final class SeriesText extends GenericChartObject implements ChartObject {

    /**
     * serialVersionUID
     */
    private static final long serialVersionUID = -3794355940075116165L;
    private final byte[] PROTOTYPE_BYTES = new byte[]{0, 0, 7, 1, 74, 0, 97, 0, 110, 0, 117, 0, 97, 0, 114, 0, 121, 0};
    protected int id = -1, cch = -1;
    private String text = "";

    public static SeriesText getPrototype(String text) {
        SeriesText st = new SeriesText();
        st.setOpcode(SERIESTEXT);
        st.setData(st.PROTOTYPE_BYTES);
        st.init();
        st.setText(text);
        return st;
    }

    public void setText(String t) {
        // create a new SeriesText value from the passed-in String
        byte[] strbytes = null;
        byte uni = 0x0;
        int lent = 0;
        try {
            strbytes = t.getBytes(WorkBookFactory.UNICODEENCODING);
            uni = 0x1;
            lent = strbytes.length / 2;
        } catch (Exception e) {
            strbytes = t.getBytes();
            lent = strbytes.length;
        }
        byte[] newbytes = new byte[strbytes.length + 4];
//		byte[] lenbytes = ByteTools.shortToLEBytes((short)strbytes.length);
        newbytes[0] = 0x0;
        newbytes[1] = 0x0;
        newbytes[2] = (byte) lent;
        newbytes[3] = uni;
        System.arraycopy(strbytes, 0, newbytes, 4, strbytes.length);
        this.setData(newbytes);
        this.text = t;
    }

    public void init() {
        super.init();
        //byte[] data = this.getData();
        int multi = 2;
        if (this.getByteAt(3) == 0x0) multi = 1;
        cch = (int) this.getByteAt(2) * multi;
        if (cch < 0) cch *= -1; // strangely it can be negative...
        try {
            byte[] namebytes = this.getBytesAt(4, cch);
            try {
                text = new String(namebytes, WorkBookFactory.UNICODEENCODING);
            } catch (UnsupportedEncodingException e) {
                Logger.logInfo("Unsupported Encoding error in SeriesText: " + e);
            }
            if ((DEBUGLEVEL > 10)) Logger.logInfo("Series Text Value: " + text);

        } catch (Exception ex) {
            Logger.logWarn("SeriesText.init failed: " + ex);
        }
        //Logger.logInfo("Initialized SeriesText: "+ text);
    }

    public String toString() {
        return this.text;
    }
}
