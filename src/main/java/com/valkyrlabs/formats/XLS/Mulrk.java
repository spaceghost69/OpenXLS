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

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import com.valkyrlabs.toolkit.ByteTools;
import com.valkyrlabs.toolkit.Logger;

/**
 * <b>Mulrk: Multiple Rk Cells (BDh)</b><br>
 * This record stores up to 256 Rk equivalents in
 * <p>
 * TODO: check compatibility with Excel2007 MAXCOLS
 * <p>
 * a space-saving format.
 * offset name size contents
 * ---
 * 4 rw 2 Row Number
 * 6 colFirst 2 Column Number of the first col of multiple Rk record
 * 8 rgrkrec var Array of 6-byte RkREC objects
 * var colLast 2 Last Column containing the RkREC object
 * </p>
 * </pre>
 *
 * @see Rk
 */

public final class Mulrk extends com.valkyrlabs.formats.XLS.XLSRecord implements Mul {
    /**
     * serialVersionUID
     */
    private static final long serialVersionUID = 1438740082267768419L;
    boolean removed = false;

    /**
     * whether this mul was removed from the SheetRecs already
     *
     * @return
     */
    public boolean removed() {
        return removed;
    }

    short colFirst;
    int colLast;
    int datalen;
    int numRkRecs = 0;
    List rkrecs;

    Mulrk() {
        super();
    }

    /**
     * populate the MULRk with its data, as well as creating
     * multiple Rk records per the Rk array.
     **/
    public void init() {
        super.init();
        int datalen = this.getData().length; // getLength();

        if (datalen <= 0) {
            if (DEBUGLEVEL > -1)
                Logger.logInfo("no data in MULRk");
        } else {
            super.initRowCol();
            short s = ByteTools.readShort(this.getByteAt(2), this.getByteAt(3));
            colFirst = s;
            col = colFirst;
            s = ByteTools.readShort(this.getByteAt(datalen - 2), this.getByteAt(datalen - 1));
            colLast = (int) s;
            // get the records data only
            datalen = datalen - 6;
            byte[] rkdatax = this.getBytesAt(4, datalen);
            numRkRecs = datalen / 6;
            // rkrecs = new Rk[numRkRecs]; Now its a vector
            rkrecs = new ArrayList(numRkRecs);
            int reccount = 0;
            int rkcol = col;
            // iterate through the rk data array and create
            // a new 6-byte Rk for each.
            for (int i = 4; i < rkdatax.length;) {
                byte[] rkd = this.getBytesAt(i, 6);
                Rk r = new Rk();
                r.init(rkd, rw, rkcol++);
                if (DEBUGLEVEL > 5)
                    Logger.logInfo(" rk@" + (rkcol - 1) + ":" + r.getStringVal());
                i += 6;
                if (reccount == numRkRecs)
                    break;
                r.setMyMul(this);
                r.setSheet(getSheet());
                r.streamer = streamer;
                rkrecs.add(r);
                reccount++;
            }
            if (DEBUGLEVEL > 5)
                Logger.logInfo("Done adding Rk recs to: " + this.getCellAddress());
        }
    }

    void deleteRk(Rk rik) {
        rkrecs.remove(rik);
    }

    void addRk(Rk rik) {
        rkrecs.add(rik);
    }

    int getColFirst() {
        return colFirst;
    }

    public List getRecs() {
        return rkrecs;
    }

    /**
     * get a handle to a specific Rk for use in updating values
     */
    Rk getRk(int rnum) {
        return (Rk) rkrecs.get(rnum);
    }

    /*
     * Changes the range of
     **/
    public Mulrk splitMulrk(int splitcol) {
        if (splitcol < colFirst || splitcol > colLast)
            return null;
        Mulrk newmul = new Mulrk();
        newmul.colFirst = (short) splitcol;
        newmul.colLast = this.colLast;
        this.colLast = splitcol - 1;
        Iterator rkr = this.getRecs().iterator();
        while (rkr.hasNext()) {
            Rk r = (Rk) rkr.next();
            if (r.getRowNumber() >= splitcol) {
                this.deleteRk(r);
                newmul.addRk(r);
            }
        }
        newmul.setOpcode(getOpcode());
        newmul.setLength(getLength());
        return newmul;
    }

    /**
     * Remove an Rk from the record along with all of the
     * folloing Rks. Returns a CompatableVector
     * of RKs that have been cut off from the Mulrk.
     * this is kinda deprecated because of the splitMulrk(),
     * but could prove to be useful later...
     */
    ArrayList removeRk(Rk rok) {
        ArrayList rez = new ArrayList();
        // set the new last col of the Mulrk
        colLast = (rok.getColNumber() - 1);
        rkrecs.remove(rok);
        int z = rkrecs.size() - 1;
        for (int i = z; i >= 0; i--) {
            Rk rec = (Rk) rkrecs.get(i);
            if (rec.getColNumber() > colLast) {
                rez.add(rec);
                rkrecs.remove(i);
            }
        }
        this.updateRks();
        return rez;
    }

    /**
     * set the row
     */
    public void setRow(int i) {
        byte[] r = ByteTools.shortToLEBytes((short) i);
        System.arraycopy(r, 0, this.getData(), 0, 2);
        rw = i;
    }

    /**
     * Update the underlying byte array for the MULRk record
     * after changes have been made to individual Rk records.
     */
    void updateRks() {
        if (this.getRecs().size() < 1) {
            this.getSheet().removeRecFromVec(this);
            return;
        }
        byte[] tmp = new byte[4];
        System.arraycopy(getData(), 0, tmp, 0, 4);
        Iterator it = this.getRecs().iterator();
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        try {
            out.write(tmp);
            // loop through the Rks and copy their bytes to the MULRk byte array.
            while (it.hasNext()) {
                out.write(((Rk) it.next()).getBytes());
            }
            out.write(ByteTools.shortToLEBytes((short) colLast));
        } catch (IOException a) {
            Logger.logInfo("parsing record continues failed: " + a);
        }
        this.setData(out.toByteArray());
    }

}