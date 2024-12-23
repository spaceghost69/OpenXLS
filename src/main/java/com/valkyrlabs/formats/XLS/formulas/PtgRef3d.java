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
package com.valkyrlabs.formats.XLS.formulas;

import com.valkyrlabs.OpenXLS.ExcelTools;
import com.valkyrlabs.OpenXLS.SheetNotFoundException;
import com.valkyrlabs.formats.XLS.BiffRec;
import com.valkyrlabs.formats.XLS.Boundsheet;
import com.valkyrlabs.formats.XLS.Externsheet;
import com.valkyrlabs.formats.XLS.Formula;
import com.valkyrlabs.formats.XLS.Name;
import com.valkyrlabs.formats.XLS.WorkBook;
import com.valkyrlabs.formats.XLS.XLSRecord;
import com.valkyrlabs.toolkit.ByteTools;
import com.valkyrlabs.toolkit.FastAddVector;
import com.valkyrlabs.toolkit.Logger;

/**
 * A BiffRec range spanning 3rd dimension of WorkSheets.
 * `
 * <p>
 * 
 * <pre>
 * offset name size contents
 * ---
 * 0 ixti 2 Index to Externsheet Sheet Record
 * 2 row 2 The row
 * 4 grCol 2 The col, or the col offset (see next table)
 *
 * the low-order 8 bytes store the col numbers. The 2 MSBs specify whether the
 * row
 * and col refs are relative or absolute.
 *
 * bits mask name content
 * ---
 * 15 8000h fRwRel =1 if the row is relative, 0 if absolute
 * 14 4000h fColRel =1 if the col is relative, 0 if absolute
 * 13-8 3F00h (reserved)
 * 7-0 00FFh col the col number or col offset (0-based)
 *
 * For 3D references, the tokens contain a negative EXTERNSHEET index,
 * indicating a reference into the own workbook.
 * The absolute value is the one-based index of the EXTERNSHEET record that
 * contains the name of the first sheet. The
 * tokens additionally contain absolute indexes of the first and last referenced
 * sheet. These indexes are independent of the
 * EXTERNSHEET record list. If the referenced sheets do not exist anymore, these
 * indexes contain the value FFFFH (3D
 * reference to a deleted sheet), and an EXTERNSHEET record with the special
 * name <04H> (own document) is used.
 * Each external reference contains the positive one-based index to an
 * EXTERNSHEET record containing the URL of the
 * external document and the name of the sheet used. The sheet index fields of
 * the tokens are not used.
 *
 * @see Ptg
 * @see Formula
 */
public class PtgRef3d extends PtgRef implements Ptg, IxtiListener {

    private static final long serialVersionUID = -441121385905948168L;
    public short ixti;
    boolean quoted = false;

    /**
     * 0x3A Reference class token: The reference address itself, independent of the
     * cell contents.
     * • 0x5A Value class token: A value (a constant, a function result, or one
     * specific value from a dereferenced cell range).
     * • 0x7A Array class token: An array of values (array of constant values, an
     * array function result, or all values of a cell range).
     */
    public PtgRef3d() {
        record = new byte[PTG_REF3D_LENGTH];
        ptgId = 0x5A; // id varies with type of token see above and setPtgType below
        record[0] = ptgId; // ""
        this.is3dRef = true;
    }

    public PtgRef3d(boolean addToRefTracker) {
        this.setUseReferenceTracker(addToRefTracker);
        ptgId = 0x5A; // TODO: id varies with type of token see above
        record[0] = ptgId; // ""
        this.is3dRef = true;

    }

    public PtgRef3d(String addr, short _ixti) {
        this();
        setLocation(addr);
        this.is3dRef = true;
    }

    public void setParentRec(XLSRecord r) {
        super.setParentRec(r);
    }

    public void addListener() {
        try {
            getParentRec().getWorkBook().getExternSheet().addPtgListener(this);
        } catch (Exception e) {
            // no need to output here. NullPointer occurs when a ref has an invalid ixti,
            // such as when a sheet was removed Worksheet exception could never really
            // happen.
        }
    }

    /**
     * @return Returns the ixti.
     */
    public short getIxti() { // only valid for 3d refs
        return ixti;
    }

    public void setIxti(short ixf) {
        if (ixti != ixf) {
            ixti = ixf;
            // this seems to be only one byte...
            if (record != null) {
                record[1] = (byte) ixf;
            }
            updateRecord();
        }
    }
    // true is relative, false is absolute

    /**
     * returns true if this PtgRef3d's ixti refers to an external sheet reference
     *
     * @return
     */
    public boolean isExternalLink() {
        try {
            return (getParentRec().getWorkBook().getExternSheet().getIsExternalLink(ixti));
        } catch (Exception e) {
            return false;
        }
    }

    public int getLength() {
        return PTG_REF3D_LENGTH;
    }

    public boolean getIsOperand() {
        return true;
    }

    public boolean getIsReference() {
        return true;
    }

    /**
     * set the Ptg Id type to one of:
     * VALUE, REFERENCE or Array
     * <br>
     * The Ptg type is important for certain
     * functions which require a specific type of operand
     */
    public void setPtgType(short type) {
        switch (type) {
            case VALUE:
                ptgId = 0x5A;
                break;
            case REFERENCE:
                ptgId = 0x3A;
                break;
            case Ptg.ARRAY:
                ptgId = 0x7A;
                break;
        }
        record[0] = ptgId;
    }

    public void init(byte[] b) {
        ptgId = b[0];
        record = b;
        populateVals();
    }

    /**
     * get the worksheet that this ref is on
     * for some reason this seems to be backwards in Ref3d
     */
    public Boundsheet getSheet(WorkBook b) {
        Boundsheet[] bsa = b.getExternSheet().getBoundSheets(ixti);
        if (bsa != null && bsa[0] == null) { // 20080303 KSC: catch error
            // try harder...
            if (parent_rec.getSheet() != null) {
                return parent_rec.getSheet(); // sheetless names belong to parent rec
            } else {
                if (b.getFactory().getDebugLevel() > 1) // 20080925 KSC
                    Logger.logErr("PtgRef3d.getSheet: Unresolved External or Deleted Sheet Reference Found"); // [BUGTRACKER
                                                                                                              // 1836]
                                                                                                              // Claritas
                                                                                                              // extenXLS22677.rec
                                                                                                              // (Deleted
                                                                                                              // Sheet/Named
                                                                                                              // Range
                                                                                                              // causes
                                                                                                              // errant
                                                                                                              // value
                                                                                                              // in B3)
                return null; // 20080805 KSC: Don't just return the 1st sheet, may be wrong, deleted, etc!
            }
        } else if (bsa == null)
            return null;
        return bsa[0];
    }

    /**
     * Throw this data into a ptgref's
     */
    public void populateVals() {
        ixti = ByteTools.readShort(record[1], record[2]);
        this.sheetname = this.getSheetName();

        rw = readRow(record[3], record[4]);
        short column = ByteTools.readShort(record[5], record[6]);
        // is the Row relative?
        fRwRel = (column & 0x8000) == 0x8000;
        // is the Column relative?
        fColRel = (column & 0x4000) == 0x4000;
        col = (short) (column & 0x3fff);
        setRelativeRowCol(); // set formulaRow/Col for relative references if necessary
        this.getIntLocation(); // sets the wholeRow and/or wholeCol flag for certain refs
        this.hashcode = super.getHashCode();
    }

    /**
     * Set the location of this PtgRef. This takes a location
     * such as "a14"
     */
    public void setLocation(String address, short ix) {
        ixti = ix;
        String[] s = ExcelTools.stripSheetNameFromRange(address);
        this.setLocation(s);
    }

    public String toString() {
        String ret = "";
        try {
            ret = getLocation();
            if ((ret.indexOf("!") == -1) && (sheetname != null)) { // prepend sheetname
                if (sheetname.indexOf(' ') == -1 && sheetname.charAt(0) != '\'') // 20081211 KSC: Sheet names with
                                                                                 // spaces must have surrounding quotes
                    ret = sheetname + "!" + ret;
                else
                    ret = "'" + sheetname + "'!" + ret;
            }
        } catch (Exception ex) {
            Logger.logErr("PtgRef3d.toString() failed", ex);
        }
        return ret;
    }

    /**
     * Change the sheet reference to the passed in boundsheet
     *
     * @see com.valkyrlabs.formats.XLS.formulas.PtgArea3d#setReferencedSheet(com.valkyrlabs.formats.XLS.Boundsheet)
     */
    public void setReferencedSheet(Boundsheet b) {
        int boundnum = b.getSheetNum();
        Externsheet xsht = b.getWorkBook().getExternSheet(true);
        // TODO: add handling for multi-sheet reference. Already handled in externsheet
        try {
            int xloc = xsht.insertLocation(boundnum, boundnum);
            setIxti((short) xloc);
            this.sheetname = null; // 20100218 KSC: RESET
            this.getSheetName();
            locax = null;
        } catch (SheetNotFoundException e) {
            Logger.logErr("Unable to set referenced sheet in PtgRef3d " + e);
        }
    }

    /**
     * Returns the location of the Ptg as a string, including sheet name
     */
    public String getLocation() {
        String ret = super.getLocation();
        if (ret.indexOf("!") == -1) { // doesn't have a sheet ref
            // NOTE: Our tests error when PtgRefs have fully qualified range syntax
            if (sheetname == null)
                sheetname = this.getSheetName();
            if (this.sheetname != null) {
                if (sheetname.equals("#REF!"))
                    return sheetname + ret;
                sheetname = qualifySheetname(sheetname);
                return sheetname + "!" + ret; // PtgRef does not have ixti
            }
        }
        return ret;
    }

    /**
     * set Ptg to parsed location
     *
     * @param loc String[] sheet1, range, sheet2, exref1, exref2
     */
    public void setLocation(String[] s) {
        if (useReferenceTracker && !getIsRefErr())
            this.getParentRec().getWorkBook().getRefTracker().removeCellRange(this);
        sheetname = null;
        if (s[0] != null) {
            sheetname = s[0];
        } else {
            try { // if not provided, assume that parent rec sheet is correct
                sheetname = this.getParentRec().getSheet().getSheetName();
            } catch (NullPointerException e) {
            }
        }
        String loc = s[1];
        if (sheetname != null) {
            loc = sheetname + "!" + loc; // loc uses quoted vers of sheet
            if (sheetname.indexOf("'") == 0) {
                sheetname = sheetname.substring(1, sheetname.length() - 1);
                quoted = true;
            }
        }
        if (sheetname != null) {
            Externsheet xsht = null;
            WorkBook b = parent_rec.getWorkBook();
            if (b == null)
                b = parent_rec.getSheet().getWorkBook();
            try {

                int boundnum = b.getWorkSheetByName(sheetname).getSheetNum();
                xsht = b.getExternSheet();
                try {
                    int xloc = xsht.insertLocation(boundnum, boundnum);
                    setIxti((short) xloc);
                } catch (Exception e) {
                    Logger.logWarn("PtgRef3d.setLocation could not update Externsheet:" + e.toString());
                }
            } catch (SheetNotFoundException e) {
                try {
                    xsht = b.getExternSheet();
                    int boundnum = xsht.getXtiReference(s[0], s[0]);
                    if (boundnum == -1) { // can't resolve
                        this.setIxti((short) xsht.insertLocation(boundnum, boundnum));
                    } else {
                        this.setIxti((short) boundnum);
                    }
                } catch (Exception ex) {
                }
            }
        }
        super.setLocation(s);
    }

    /**
     * Set Location can take either a local page address (ie A54) or
     * a reference to a page and location(ie Sheet2!A22). It then changes
     * the location reference of the Ptg.
     *
     * @see com.valkyrlabs.formats.XLS.formulas.Ptg#setLocation(java.lang.String)
     */
    public void setLocation(String address) {
        String[] s = ExcelTools.stripSheetNameFromRange(address);
        setLocation(s);
    }

    /**
     * Updates the record bytes so it can be pulled back out.
     */
    public void updateRecord() {
        byte[] tmp = new byte[PTG_REF3D_LENGTH];
        tmp[0] = record[0];
        byte[] ix = ByteTools.shortToLEBytes(ixti);
        System.arraycopy(ix, 0, tmp, 1, 2);
        byte[] brow = ByteTools.cLongToLEBytes(rw);
        System.arraycopy(brow, 0, tmp, 3, 2);
        if (fRwRel) {
            col = (short) (0x8000 | col);
        }
        if (fColRel) {
            col = (short) (0x4000 | col);
        }
        byte[] bcol = ByteTools.cLongToLEBytes(col);
        if (col == -1) { // KSC: what excel expects
            bcol[1] = 0;
        }
        System.arraycopy(bcol, 0, tmp, 5, 2);
        record = tmp;
        if (parent_rec != null) {
            if (this.parent_rec instanceof Formula)
                ((Formula) this.parent_rec).updateRecord();
            else if (this.parent_rec instanceof Name)
                ((Name) this.parent_rec).updatePtgs();
        }

        col = (short) col & 0x3FFF; // get lower 14 bits which represent the actual column;
    }

    public Boundsheet getSheet() {
        if (parent_rec != null) {
            WorkBook wb = parent_rec.getWorkBook();
            if (wb != null && wb.getExternSheet() != null) {
                Boundsheet[] bsa = wb.getExternSheet().getBoundSheets(this.ixti);
                if (bsa == null || bsa[0] == null) {// 20080303 KSC: Catch Unresolved External refs
                    if (parent_rec instanceof Formula)
                        Logger.logErr("PtgRef3d.getSheet: Unresolved External Worksheet in Formula "
                                + parent_rec.getCellAddressWithSheet());
                    else if (parent_rec instanceof Name)
                        Logger.logErr("PtgRef3d.getSheet: Unresolved External Worksheet in Name "
                                + ((Name) parent_rec).getName());
                    else
                        Logger.logErr("PtgRef3d.getSheet: Unresolved External Worksheet for "
                                + parent_rec.getCellAddressWithSheet());
                    return null;
                }
                return bsa[0];
            }
        }
        return null;
    }

    /**
     * return the sheet name for this 3d reference
     */
    public String getSheetName() {
        if (this.sheetname == null) {
            if (parent_rec != null) {
                WorkBook wb = parent_rec.getWorkBook();
                if (wb != null && wb.getExternSheet() != null) { // 20080306 KSC: new way is to get sheet names rather
                                                                 // than sheets as can be external refs
                    String[] sheets = wb.getExternSheet().getBoundSheetNames(this.ixti);
                    if (sheets != null && sheets[0] != null)
                        sheetname = sheets[0];
                }
            }
        }
        return sheetname;
    }

    /**
     * @return Returns the refCell.
     */
    public BiffRec[] getRefCells() {
        if (sheetname == null)
            sheetname = this.getSheetName();
        refCell = super.getRefCells();
        return refCell;
    }

    /**
     * PtgRef's have no sub-compnents
     */
    public Ptg[] getComponents() {
        return null; // only one
    }

    /**
     * return the ptg components for a certain column within a ptgArea()
     *
     * @param colNum
     * @return all Ptg's within colNum
     */
    public Ptg[] getColComponents(int colNum) {
        FastAddVector v = new FastAddVector();
        int[] x = this.getIntLocation();
        if (x[1] == colNum)
            v.add(this);
        PtgRef[] pref = new PtgRef[v.size()];
        v.toArray(pref);
        return pref;
    }
}