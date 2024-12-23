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

import com.valkyrlabs.OpenXLS.Cell;
import com.valkyrlabs.OpenXLS.ExcelTools;
import com.valkyrlabs.OpenXLS.SheetNotFoundException;
import com.valkyrlabs.formats.XLS.BiffRec;
import com.valkyrlabs.formats.XLS.Boundsheet;
import com.valkyrlabs.formats.XLS.Dbcell;
import com.valkyrlabs.formats.XLS.ExpressionParser;
import com.valkyrlabs.formats.XLS.Formula;
import com.valkyrlabs.formats.XLS.Name;
import com.valkyrlabs.formats.XLS.Row;
import com.valkyrlabs.formats.XLS.WorkBook;
import com.valkyrlabs.formats.XLS.XLSRecord;
import com.valkyrlabs.formats.cellformat.CellFormatFactory;
import com.valkyrlabs.toolkit.ByteTools;
import com.valkyrlabs.toolkit.Logger;
import com.valkyrlabs.toolkit.StringTool;


/**
 * ptgRef is a reference to a single cell.  It contains row and
 * column information, plus a grbit to determine whether these
 * values are relative or absolute.  This grbit is, stupidly, but expectedly,
 * encoded within the column value.
 * <pre>
 * Offset      Name        Size    Contents
 * ----------------------------------------------------
 * 0           rw          2       the row
 * 2           grbitCol    2       (see following table)
 *
 * Only the low-order 14 bits specify the Col, the other bits specify
 * relative vs absolute for both the col or the row.
 *
 * Bits        Mask        Name    Contents
 * -----------------------------------------------------
 * 15          8000h       fRwRel  =1 if row offset relative,
 * =0 if otherwise
 * 14          4000h       fColRel =1 if col offset relative,
 * =0 if otherwise
 * 13-0        3FFFh       col     Ordinal column offset or number
 * </pre>
 *
 * @see WorkBook
 * @see Boundsheet
 * @see Dbcell
 * @see Row
 * @see Cell
 * @see XLSRecord
 */
public class PtgRef extends GenericPtg implements Ptg {
    /**
     *
     */
    private static final long serialVersionUID = -7776520933300730470L;
    //    String locstrax = null;
    public long hashcode = -1L;
    public boolean wholeRow = false, wholeCol = false;    // denotes a range which spans the entire row or column, a shorthand for checking end col or row # as this will vary between excel versions
    protected int rw;
    // TODO: We actually are talking about 2 different notions of relativity:
    // 1- Relativity based on shared formula parent formula row/col
    // 2- Relative/Absolute in terms of row movement ($'s mean reference is ABSOLUTE)
    // we are combining the two concepts erroneously
    protected boolean fRwRel = true;  //true is relative, false is absolute (=$'s)
    protected boolean fColRel = true;
    protected int col;
    protected boolean is3dRef = false;
    protected int formulaRow;
    protected int formulaCol;
    protected String sheetname = null;

    protected int externalLink1 = 0;
    protected int externalLink2 = 0;
    protected boolean useReferenceTracker = true;
    BiffRec[] refCell = new BiffRec[1];
    String locax = null;
    private String cachedLocation = null;

    public PtgRef(int[] rowcol, XLSRecord x, boolean useRefTracker) {
        this();
        setParentRec(x);
        this.useReferenceTracker = useRefTracker;
        setLocation(rowcol);
        updateRecord();
    }

    /**
     * This constructor is for programmatic creation of Ptg's
     * in this case we do not have the ptgid, just the refereced location
     */
    public PtgRef(String location, XLSRecord x, boolean utilizeRefTracker) {
        this.setUseReferenceTracker(utilizeRefTracker);
        ptgId = 0x44;  //0x24; defaulting to value operand
        record = new byte[5];
        record[0] = ptgId;
        setParentRec(x);    // MUST set before setLocation also sets formulaRow ...
        this.setLocation(location);
        this.setIsWholeRowCol();
        if (useReferenceTracker)
            addToRefTracker();
    }

    /**
     * This constructor is for programmatic creation of Ptg's
     * in this case we do not have the ptgid, just the refereced location
     * <p>
     * this version sets the value of useReferenceTracker to avoid multiple entries due to area parent
     */
    public PtgRef(byte[] bin, XLSRecord x, boolean utilizeRefTracker) {
        this();
        this.setUseReferenceTracker(utilizeRefTracker);
        setParentRec(x);    //MUST DO BEFORE INIT ... also sets formulaRow ...
        init(bin);
        if (useReferenceTracker)
            addToRefTracker(); // TODO: check subreference issue (if it's not a 'real' ptg)
    }

    /**
     * default constructor
     */
    public PtgRef() {
        // 24H (tRefR), 44H (tRefV), 64H (tRefA)
        ptgId = 0x44;  // default to value operand
        record = new byte[5];
        record[0] = ptgId;
    }

    /**
     * given an address string, parse and assign to the appropriate PtgRef-type object
     * <br>#REF! 's return either PtgRefErr or PtgRefErr3d
     * <br>Ranges return either PtgArea or PtgArea3d
     * <br>Single addresses return either PtgRef or PtgRef3d
     * <br>NOTE: This method does not extract names embedded within the address string
     *
     * @param address
     * @param parent  parent record to assign the ptg to
     * @return
     */
    public static Ptg createPtgRefFromString(String address, XLSRecord parent) {
        try {
            String[] s = ExcelTools.stripSheetNameFromRange(address);
            String sh1 = s[0];
            String range = s[1];
            Ptg ptg;
            if (range == null || range.equals("#REF!") || (sh1 != null && sh1.equals("#REF"))) {
                if (sh1 != null) {
                    PtgRefErr3d pe3 = new PtgRefErr3d();
                    pe3.setParentRec(parent);
                    pe3.setLocation(s);
                    return pe3;
                } else {
                    PtgRefErr pe = new PtgRefErr();
                    pe.setParentRec(parent);
                    pe.setLocation(s);
                    return pe;
                }
            }
            WorkBook bk = parent.getWorkBook();


            String sht = "((?:\\\\?+.)*?!)?+";
            String rangeMatch = "(.*(:).*){2,}?";    //matches 2 or more range ops (:'s)
            String opMatch = "(.*([ ,]).*)+";        //matches union or isect op	( " " or ,)
            String m = sht + "((" + opMatch + ")|(" + rangeMatch + "))";
            // is address a complex range??
            if (address.matches(m) || range.indexOf("(") > -1) {
                //NOTE: this can be a MemFunc OR a MemArea --
                // PtgMemFunc= a NON-CONSTANT cell address, cell range address or cell range list
                // Whenever one operand of the reference subexpression is a function, a defined name, a 3D
                // reference, or an external reference (and no error occurs), a PtgMemFunc token is used.
                // PtgMemArea= constant cell address, cell range address, or cell range list on the same sheet
                PtgMemFunc pmf = new PtgMemFunc();
                pmf.setParentRec(parent);
                pmf.setLocation(address);    // TODO HANDLE FUNCTION MEMFUNCS ALA OFFSET(x,y,0):OFFSET(x,y,0)
                ptg = pmf;
            } else if (range.indexOf(":") > 0) { // it's a range, either PtgRef3d or PtgArea3d
                String[] ops = StringTool.getTokensUsingDelim(range, ":");
                if ((bk.getName(ops[0]) != null || bk.getName(ops[1]) != null)) {
                    PtgMemFunc pmf = new PtgMemFunc();
                    pmf.setParentRec(parent);
                    pmf.setLocation(address);
                    ptg = pmf;
                } else if (sh1 != null) {
                    int[] rc = ExcelTools.getRowColFromString(ops[0]);    // see if a wholerow/wholecol ref
                    if (!(ops[0].equals(ops[1]) && rc[0] != -1 && rc[1] != -1)) {
                        PtgArea3d pta = new PtgArea3d();
                        pta.setParentRec(parent);
                        pta.setLocation(s);
                        ptg = pta;
                    } else {
                        ptg = new PtgRef3d();
                        ((PtgRef3d) ptg).setPtgType(REFERENCE);
                        ptg.setParentRec(parent);
                        ((PtgRef) ptg).setUseReferenceTracker(false);
                        ((PtgRef3d) ptg).setLocation(s);
                        ((PtgRef) ptg).setUseReferenceTracker(true);
                        ((PtgRef3d) ptg).addToRefTracker();
                    }
                } else {
                    PtgArea pa = new PtgArea();
                    pa.setParentRec(parent);
                    pa.setUseReferenceTracker(false);
                    pa.setLocation(s);
                    pa.setUseReferenceTracker(true);
                    pa.addToRefTracker();
                    ptg = pa;
                }
            } else { // it's a single ref NOT a range e.g. Sheet1!A1
                if (sh1 != null) {
                    ptg = new PtgRef3d();
                    ((PtgRef3d) ptg).setPtgType(REFERENCE);
                    ptg.setParentRec(parent);
                    ((PtgRef) ptg).setUseReferenceTracker(false);
                    ((PtgRef3d) ptg).setLocation(s);
                    ((PtgRef) ptg).setUseReferenceTracker(true);
                    ((PtgRef3d) ptg).addToRefTracker();
                } else {
                    PtgRef pr = new PtgRef();
                    pr.setParentRec(parent);
                    pr.setUseReferenceTracker(false);
                    pr.setLocation(s);
                    pr.setUseReferenceTracker(true);
                    pr.addToRefTracker();
                    ptg = pr;
                }
            }
            return ptg;
        } catch (Exception e) {    // any error in parsing return a referr -- makes sense!!!
            PtgRefErr3d pe3 = new PtgRefErr3d();
            pe3.setParentRec(parent);
            return pe3;
        }
    }

    public static long getHashCode(int row, int col) {
        return col + ((row + 1) * MAXCOLS);
    }

    public boolean equals(Object ob) {
        return ob.hashCode() == this.hashCode();
    }

    public boolean getIsWholeRow() {
        return wholeRow;
    }

    public void setIsWholeRow(boolean b) {
        wholeRow = b;
    }

    public boolean getIsWholeCol() {
        return wholeCol;
    }


    /**
     * @param id
     */

    public void setIsWholeCol(boolean b) {
        wholeCol = b;
    }

    public boolean getIsRefErr() {
        return false;
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
     * <br>The Ptg type is important for certain
     * functions which require a specific type of operand
     */
    public void setPtgType(short type) {
        switch (type) {
            case VALUE:
                ptgId = 0x44;
                break;
            case REFERENCE:
                ptgId = 0x24;
                break;
            case Ptg.ARRAY:
                ptgId = 0x64;
                break;
        }
        record[0] = ptgId;
    }

    public void init(byte[] b) {
        ptgId = b[0];
        record = b;
        this.populateVals();
    }

    public boolean getUseReferenceTracker() {
        return useReferenceTracker;
    }

    /**
     * Ptgs upkeep their mapping in reference tracker, however, some ptgs
     * are components of other Ptgs, such as individual ptg cells in a PtgArea.  These
     * should not be stored in the RT.
     */
    public void setUseReferenceTracker(boolean b) {
        useReferenceTracker = b;
    }

    /**
     * parse all the values out of the byte array and
     * populate the classes values
     */
    protected void populateVals() {
        rw = readRow(record[1], record[2]);
        short column = ByteTools.readShort(record[3], record[4]);
        // is the Row relative?
        fRwRel = (column & 0x8000) == 0x8000;
        // is the Column relative?
        fColRel = (column & 0x4000) == 0x4000;
        col = (short) (column & 0x3fff);
        setRelativeRowCol();  // set formulaRow/Col for relative references if necessary
        this.setIsWholeRowCol();
        hashcode = getHashCode();
    }

    public boolean is3dRef() {
        return is3dRef;
    }

    /**
     * return the human-readable String representation of
     * this ptg -- if applicable
     */
    public String getString() {
        return this.getLocation();
    }

    /**
     * returns the String address of this ptg including sheet reference
     *
     * @return
     */
    public String getLocationWithSheet() {
        String ret = getString();

        // AI PtgRefs do not have location info
        if ((ret == null) && (parent_rec.getOpcode() == XLSRecord.AI))
            return parent_rec.toString();

        if (ret == null)
            return "";

        if (ret.indexOf("!") > -1)
            return ret;

        ret = sheetname + "!" + ret;

        return ret;
    }

    public String toString() {
        return getString();
    }

    /**
     * returns the row/col ints for the ref
     *
     * @return
     */
    public int[] getRowCol() {
        int[] ret = {this.rw, this.col};
        if (this.rw < 0) {// if row truly references MAXROWS_BIFF8 comes out -
            ret[0] = MAXROWS_BIFF8;
            this.wholeCol = true;
        }
        return ret;
    }

    /**
     * Returns the location of the Ptg as a string (ie c4)
     *
     * @see com.valkyrlabs.formats.XLS.formulas.GenericPtg#getLocation()
     */
    public String getLocation() {
        if (locax != null)//cache
            return locax;

        int[] adjusted = this.getIntLocation();
        String s;
        if (this.wholeCol) {
            s = (fColRel ? "" : "$") + ExcelTools.getAlphaVal(adjusted[1]);
        } else if (this.wholeRow) {
            s = (fRwRel ? "" : "$") + (adjusted[0] + 1);
        } else {
            if (rw < 0 || col < 0)
                return new PtgRefErr().toString();

            s = (fColRel ? "" : "$") + ExcelTools.getAlphaVal(adjusted[1]) +
                    (fRwRel ? "" : "$") + (adjusted[0] + 1);
        }
        locax = s;
        return locax;
    }

    /**
     * Set the location of this PtgRef.  This takes a location
     * such as "a14",   also can take a absolute location, such as $A14
     */
    public void setLocation(String address) {
        locax = null;
        refCell = null;
        if (record != null) {
            String[] s = ExcelTools.stripSheetNameFromRange(address);
            setLocation(s);
            locax = s[1];
        } else {
            Logger.logWarn("PtgRef.setLocation() failed: NO record data: " + address);
        }
    }

    /**
     * set Ptg to parsed location
     *
     * @param loc String[] sheet1, range, sheet2, exref1, exref2
     */
    public void setLocation(String[] loc) {
        if (useReferenceTracker)
            removeFromRefTracker();
        locax = null;
        sheetname = loc[0];
        String addr = loc[1];
        cachedLocation = addr;
        fRwRel = true;
        fColRel = true;
        if (addr.indexOf("$") == -1) {    // both row and col are relative refs, meaning moves/copies will change ref
            // relative link
            if (!addr.equals("#REF!") && !addr.equals("")) {
                int[] res = ExcelTools.getRowColFromString(addr);
                col = res[1];
                rw = res[0];
            } else {
                col = -1;
                rw = -1;
            }
        } else {
            // absolute reference
            if (addr.substring(0, 1).equalsIgnoreCase("$")) {
                fColRel = false;
                addr = addr.substring(1);
            }
            if (addr.indexOf("$") != -1) {
                fRwRel = false;
                addr = StringTool.strip(addr, "$");
            }
            int[] res = null;
            try {
                res = ExcelTools.getRowColFromString(addr);
                col = res[1];
                rw = res[0];
                if (col == -1 || rw == -1) {    // if wholerow or wholecol, must be absolute
                    fColRel = false;
                    fRwRel = false;
                }
            } catch (IllegalArgumentException ie) {    //is it a wholerow/wholecol issue?
                if (Character.isDigit(addr.charAt(0))) { //assume wholecol ref
                    col = MAXCOLS_BIFF8 - 1;
                    rw = Integer.valueOf(addr).intValue() - 1;
                    fColRel = false;
                    fRwRel = false;
                } else { //wholerow ref?
                    rw = -1;
                    col = ExcelTools.getIntVal(addr);
                    fColRel = false;
                    fRwRel = false;
                }
            }
        }
        if (col == -1)
            wholeRow = true;
        if (rw == -1)
            wholeCol = true;
        this.setIsWholeRowCol();
        this.updateRecord();
        hashcode = getHashCode();
        // trap OOXML external reference link, if any
        if (loc[3] != null)
            externalLink1 = Integer.valueOf(loc[3].substring(1, loc[3].length() - 1)).intValue();
        if (loc[4] != null)
            externalLink2 = Integer.valueOf(loc[4].substring(1, loc[4].length() - 1)).intValue();
        if (useReferenceTracker) {
            if (!getIsRefErr() && !this.getIsWholeCol() && !this.getIsWholeRow())
                addToRefTracker();
        }
    }

    /**
     * Set the location of this PtgRef.  This takes a location
     * such as {1,2}
     */
    public void setLocation(int[] rowcol) {
        locax = null;
        cachedLocation = null;
        if (record != null) {
            if (useReferenceTracker)
                removeFromRefTracker();
            rw = rowcol[0];
            col = rowcol[1];
            fRwRel = true;    // default
            fColRel = true;
            this.updateRecord();
            hashcode = getHashCode();
            if (useReferenceTracker)
                addToRefTracker();
        } else {
            Logger.logWarn("PtgRef.setLocation() failed: NO record data: " + rowcol.toString());
        }
    }

    /**
     * Get the location of this ptgRef as an int array {row, col}.  0 based
     */
    public int[] getIntLocation() {

        this.setIsWholeRowCol();
        int rowNew = rw;
        int colNew = col;
        try {
            boolean isExcel2007 = this.parent_rec.getWorkBook().getIsExcel2007();
            if (fRwRel) {  // the row is a relative location
                rowNew += formulaRow;
            }
            if (fColRel) {  // the column is a relative location
                colNew += formulaCol;
            }
            if (wholeRow) {
                if (!isExcel2007)
                    colNew = MAXCOLS_BIFF8;
                else
                    colNew = MAXCOLS;
            }
            if (wholeCol) {
                if (isExcel2007)
                    rowNew = MAXROWS - 1;
                else
                    rowNew = MAXROWS_BIFF8 - 1;
            }

        } catch (NullPointerException e) {
        }
        return new int[]{rowNew, colNew};
    }

    /**
     * Get the location of this ptgRef as an int array {row, col}.  0 based
     * NOTE: this version of getIntLocation returns the actual or real coordinates
     * This may be different from getIntLocation when rw designates MAXROWS - in these cases,
     * this method will return real max rows
     */
    public int[] getRealIntLocation() {
        int rowNew = rw;
        int colNew = col;
        if (fRwRel) {  // the row is a relative location
            rowNew += formulaRow;
        }
        if (fColRel) {  // the column is a relative location
            colNew += formulaCol;
        }

        if (wholeCol || rowNew < 0)
            try {
                if (rowNew < 0)
                    wholeCol = true;
                rowNew = this.getParentRec().getSheet().getMaxRow();
            } catch (Exception e) {
            }

        if (wholeRow || colNew >= MAXCOLS)
            try {
                colNew = this.getParentRec().getSheet().getMaxCol();
            } catch (Exception e) {
            }
        int[] ret = {rowNew, colNew};
        return ret;
    }

    /**
     * Get the worksheet name this ptgref refers to
     *
     * @throws SheetNotFoundException{
     */
    public String getSheetName()
            throws SheetNotFoundException{

        if (locax != null) { // reference on different sheet than parent
            if (locax.indexOf("!") > -1) {
                sheetname = locax.substring(0, locax.indexOf("!"));
            }
        }
        if (sheetname == null &&
                parent_rec != null) {
            if (parent_rec.getSheet() != null) {
                sheetname = parent_rec.getSheet().getSheetName();
            }
        }

        if (sheetname == null) {
            return ""; // no sheetname
        } else {

            //handle external references (OOXML-specific)
            if (externalLink1 > 0) {
                if (sheetname.charAt(0) == '\'')
                    sheetname = sheetname.substring(1, sheetname.length() - 1);
                sheetname = "[" + externalLink1 + "]" + sheetname;
            }
            sheetname = qualifySheetname(sheetname);

        }
        return sheetname;
    }

    /**
     * sets the sheetname for this
     *
     * @param sheetname
     */
    public void setSheetName(String sheetname) {
        this.sheetname = sheetname;
    }

    /**
     * Clears the location cache when needed
     */
    public void clearLocationCache() {
        locax = null;
    }

    /**
     * Does this ref reference an entire row (ie $1);
     *
     * @return
     */
    private boolean referencesEntireRow() {
        boolean isExcel2007 = this.parent_rec.getWorkBook().getIsExcel2007();
        int colNew = col;
        if (fColRel) {  // the row is a relative location
            colNew += formulaRow;
        }
        if (colNew < 0) {   // have to assume that it's a wholeRow even if 2007
            return true;
        } else if (colNew >= MAXCOLS_BIFF8 - 1 && !isExcel2007) {
            return true;
        }
        if (this.cachedLocation != null && isExcel2007) {
            return this.locationStringReferencesEntireRow();
        }
        // This is unfortunately a bit of a hack due to biff 8 incompatibilies
        return colNew == MAXCOLS_BIFF8 - 1 && isExcel2007;

    }

    /**
     * Check if the cached string location referrs to a full row
     *
     * @return
     */
    private boolean locationStringReferencesEntireRow() {
        if (this.cachedLocation != null) {
            int[] res = ExcelTools.getRowColFromString(cachedLocation);
            return res[1] < 0;
        }
        return false;
    }

    /**
     * Does this ref reference an entire col (ie $A);
     *
     * @return
     */
    private boolean referencesEntireCol() {
        int rowNew = rw;
        boolean isExcel2007 = this.parent_rec.getWorkBook().getIsExcel2007();
        if (fRwRel) {  // the row is a relative location
            rowNew += formulaRow;
        }
        if (rowNew < 0) {
            return true;
        } else if (rowNew >= MAXROWS_BIFF8 - 1 && !isExcel2007) {
            rowNew = -1;
            return true;
        }
        return false;
    }

    /**
     * Inspects the record to determin if it references whole
     * rows or columns and sets the values as required.
     *
     * @return
     */
    protected void setIsWholeRowCol() {
        this.wholeCol = referencesEntireCol();
        this.wholeRow = referencesEntireRow();
    }

    /**
     * set the location of this PtgRef
     *
     * @param rowcol  int[] rowcol
     * @param bRowRel true if row is relative (i.e. A1 not A$1)
     * @param bColRel true if col is relative (i.e. A1 not $A1)
     */
    public void setLocation(int[] rowcol, boolean bRowRel, boolean bColRel) {
        locax = null;
        cachedLocation = null;
        if (record != null) {
            if (useReferenceTracker) removeFromRefTracker();
            rw = rowcol[0];
            col = rowcol[1];
            fRwRel = bRowRel;
            fColRel = bColRel;
            this.updateRecord();
            if (useReferenceTracker)
                addToRefTracker();
        } else {
            Logger.logWarn("PtgRef.setLocation() failed: NO record data: " + rowcol.toString());
        }
    }

    /**
     * Updates the record bytes so it can be pulled back out.
     */
    public void updateRecord() {
        byte[] tmp = new byte[5];
        tmp[0] = record[0];
        byte[] brow = ByteTools.cLongToLEBytes(rw);
        System.arraycopy(brow, 0, tmp, 1, 2);
        if (fRwRel) {
            col = (short) (0x8000 | col);
        }
        if (fColRel) {
            col = (short) (0x4000 | col);
        }
        byte[] bcol = ByteTools.cLongToLEBytes(col);
        if (col == -1) {    // KSC: what excel expects
            bcol[1] = 0;
        }
        System.arraycopy(bcol, 0, tmp, 3, 2);

        record = tmp;
        if (parent_rec != null) {
            if (this.parent_rec instanceof Formula)
                ((Formula) this.parent_rec).updateRecord();
            else if (this.parent_rec instanceof Name)
                ((Name) this.parent_rec).updatePtgs();
        }

        col = (short) col & 0x3FFF;    //get lower 14 bits which represent the actual column;
    }

    public int getLength() {
        return PTG_REF_LENGTH;
    }

    /**
     * return truth of "reference is blank"
     *
     * @return
     */
    public boolean isBlank() {
        getRefCells();
        return (refCell[0] == null || ((XLSRecord) refCell[0]).isBlank);//getOpcode()==BLANK);
    }

    /**
     * returns the value of the cell refereced by the PtgRef
     */
    public Object getValue() {
        getRefCells();
        Object retValue = null;
        if (refCell[0] != null) {
            if (refCell[0].getFormulaRec() != null) {
                Formula f = refCell[0].getFormulaRec();
                retValue = f.calculateFormula();
                return retValue;
            } else {
                if (refCell[0].getDataType().equals("Float")) {
                    retValue = new Double(refCell[0].getDblVal());
                    return retValue;
                } else {
                    retValue = refCell[0].getInternalVal();
                    return retValue;
                }
            }
        } else {
            try {
                if (!this.parent_rec.getSheet().getWindow2().getShowZeroValues())
                    return null;
            } catch (NullPointerException e) {
                // assume zero, which the vast majority of cases are
            }
            return Integer.valueOf(0);
        }
    }

    /**
     * returns the value of the ptg formatted via the underlying cell's number format
     *
     * @return String underlying cell value formatted via cell's format pattern
     */
    public String getFormattedValue() {
        getRefCells();
        Object retValue = null;
        BiffRec cell = refCell[0];

        if (cell != null) {
            if (cell.getFormulaRec() != null) {
                Formula f = cell.getFormulaRec();
                retValue = f.calculateFormula();
            } else {
                if (cell.getDataType().equals("Float")) {
                    retValue = new Double(cell.getDblVal());
                } else {
                    retValue = cell.getInternalVal();
                }
            }
            return CellFormatFactory.fromPatternString(
                    cell.getXfRec().getFormatPattern()).format(retValue);
        } else {
            try {
                if (!this.parent_rec.getSheet().getWindow2().getShowZeroValues())
                    return "";
            } catch (NullPointerException e) {
                // assume zero, which the vast majority of cases are
            }
            return "0";
        }
    }

    /**
     * @return Returns the refCell.
     */
    public BiffRec[] getRefCells() {
        refCell = new BiffRec[1];
        try {
            Boundsheet bs = null;
            if ((sheetname != null) && (parent_rec != null)) {
                bs = this.parent_rec.getWorkBook().getWorkSheetByName(sheetname);
            } else if (parent_rec != null) {
                bs = parent_rec.getSheet();
            }
            refCell[0] = bs.getCell(rw, col);
        } catch (Exception ex) {
        }
        return refCell;
    }

    public boolean changeLocation(String newLoc, Formula f) {
        locax = null;
        Ptg ptg = null;
        int z = -1;
        try {
            z = ExpressionParser.getExpressionLocByPtg(this, f.getExpression());
            ptg = (Ptg) f.getExpression().get(z);
        } catch (Exception e) {

        }
        String unstripped = newLoc;

        if (newLoc.indexOf("!") > -1) newLoc = newLoc.substring(newLoc.indexOf("!") + 1);
        if (unstripped.indexOf(":") > 0) { // then either PtgRef3d or PtgArea
            if (unstripped.indexOf("!") > unstripped.indexOf(":")) { // than it's a PtgRef3d or PtgArea3d
                if (unstripped.indexOf(":") != unstripped.lastIndexOf(":")) {    // it's a PtgArea3d ala Sheet1:Sheet3!A1:D1
                    PtgArea3d pta3 = new PtgArea3d();
                    pta3.setLocation(unstripped);
                    ptg = pta3;
                } else { // it's a PtgRef3d ala Sheet1!Sheet3:A1
                    PtgRef3d prd = new PtgRef3d();
                    prd.setParentRec(f);
                    prd.setLocation(unstripped);
                    ptg = prd;
                }
                // no sheet ref ...
            } else {    // it's a PtgArea3d (according to Excel's Ai recs !!)
                PtgArea pta = new PtgArea();
                pta.setParentRec(f);
                pta.setLocation(unstripped);
                ptg = pta;
            }
        } else if (ptg != null) {
            // it's a single location
            if (!unstripped.equals("")) {
                ptg.setParentRec(f);
                ptg.setLocation(unstripped);
            } else {
                ptg = new PtgRef3d();
                ptg.setParentRec(f);
            }
        } else { // ptg is null, create a new one
            ptg = new PtgRef();
            ptg.setParentRec(f);
            ptg.setLocation(unstripped);
        }
        if (z != -1)
            f.getExpression().set(z, ptg);    // update expression with new Ptg
        else
            f.getExpression().add(ptg);
        return true;
    }

    public void setParentRec(XLSRecord f) {
        parent_rec = f;
        setRelativeRowCol(); // trap formulaRow, Col for relative PtgRefs
    }

    /**
     * removes this reference from the tracker...
     * <p>
     * used mostly when we've updated the ref and want
     * to re-register it.
     */
    public void removeFromRefTracker() {
        try {
            if (parent_rec != null) {
                parent_rec.getWorkBook().getRefTracker().removeCellRange(this);
                if (parent_rec.getOpcode() == FORMULA) ((Formula) parent_rec).setCachedValue(null);
            }
        } catch (Exception ex) {
            // no need to error here, sometimes this is called before its in Logger.logErr("PtgRef.removeFromRefTracker() failed.", ex);
        }
    }

    /**
     * add this reference to the ReferenceTracker... this
     * is crucial if we are to update this Ptg when cells
     * are changed or added...
     */
    public void addToRefTracker() {
        //Logger.logInfo("Adding :" + this.toString() + " to tracker");
        try {
            if (parent_rec != null)
                parent_rec.getWorkBook().getRefTracker().addCellRange(this);
        } catch (Exception ex) {
            Logger.logErr("PtgRef.addToRefTracker() failed.", ex);
        }
    }

    /**
     * update existing tracked ptg with new parent in reference tracker
     *
     * @param parent
     */
    public void updateInRefTracker(XLSRecord parent) {
        try {
            if (parent != null)
                parent.getWorkBook().getRefTracker().updateInRefTracker(this, parent);
        } catch (Exception ex) {
            Logger.logErr("updateInRefTracker() failed.", ex);
        }
    }

    /**
     * set the formulaRow and formulaCol for relatively-referenced PtgRefs
     */
    public void setRelativeRowCol() {
        if (fRwRel || fColRel) {
            short opc = 0;
            if (parent_rec != null)
                opc = parent_rec.getOpcode();
            // protocol for shared formulas, conditional formatting, data validity and defined names only (type B cell addresses!)
            if (opc == SHRFMLA || opc == DVAL) {
                this.formulaRow = parent_rec.getRowNumber();
                this.formulaCol = parent_rec.getColNumber();
            }
        }
    }

    /**
     * set this Ptg to an External Location - used when copying a sheet from another workbook
     *
     * @param f parent formula rec
     */
    public void setExternalReference(String externalWorkbook) {
        if (this instanceof PtgArea3d) {
            PtgArea3d ptg = (PtgArea3d) this;
            WorkBook b = parent_rec.getWorkBook();
            if (b == null)
                b = parent_rec.getSheet().getWorkBook();
            short ixti = b.getExternSheet().addExternalSheetRef(externalWorkbook, ptg.getSheetName());        //20080714 KSC: May not reflect external reference!  this.sheetname);
            ptg.setIxti(ixti);
            if (ptg.firstPtg != null) { // it's not a Ref3d
                ptg.firstPtg.updateRecord();
                ptg.lastPtg.updateRecord();
            }
            ptg.updateRecord();
        } else if (this instanceof PtgRef3d) {
            WorkBook b = parent_rec.getWorkBook();
            PtgRef3d pr = (PtgRef3d) this;
            if (b == null)
                b = parent_rec.getSheet().getWorkBook();
            short ixti = b.getExternSheet().addExternalSheetRef(externalWorkbook, pr.getSheetName());        //20080714 KSC: May not reflect external reference!  this.sheetname);
            pr.setIxti(ixti);
        } else { // TODO: convert to ref3d?
            Logger.logWarn("PtgRef.setExternalReference: unable to convert ref");
        }
    }

    public boolean isRowRel() {
        return fRwRel;
    }

    /**
     * sets the row to be relative (relative is true) or absolute (relative is false)
     * <br>absolute references do not shift upon row inserts or deletes
     *
     * @param boolean relative
     */
    public void setRowRel(boolean relative) {
        if (fRwRel != relative) {
            locax = null;
            fRwRel = relative;
            updateRecord();
        }
    }

    public boolean isColRel() {
        return fColRel;
    }

    /**
     * sets the column to be relative (relative is true) or absolute (relative is false)
     * <br>absolute references do not shift upon column inserts or deletes
     *
     * @param boolean relative
     */
    public void setColRel(boolean relative) {
        if (fColRel != relative) {
            locax = null;
            fColRel = relative;
            updateRecord();
        }
    }

    public void setArrayTypeRef() {
        byte b = (byte) ((record[0] | 0x60));
        record[0] = b;
    }

    /**
     * uniquely identifies a row/col
     * to unencrypt:
     * col= hashcode%maxcols
     * row= hashcode/maxcols -1
     */
    protected long getHashCode() {
        if (rw >= 0)
            return col + ((rw + 1) * MAXCOLS);
        else
            return col + ((MAXROWS - rw + 1) * MAXCOLS);
    }

    /**
     * clear out object references in prep for closing workbook
     */
    public void close() {
        if (useReferenceTracker) removeFromRefTracker();
        useReferenceTracker = false;
        super.close();
        if (refCell != null && refCell.length > 0 && refCell[0] != null) // clear out object references
            ((XLSRecord) refCell[0]).close();
        refCell = null;
    }

}