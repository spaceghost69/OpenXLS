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

import java.util.Stack;

import com.valkyrlabs.OpenXLS.ExcelTools;
import com.valkyrlabs.formats.XLS.formulas.FormulaCalculator;
import com.valkyrlabs.formats.XLS.formulas.FormulaParser;
import com.valkyrlabs.formats.XLS.formulas.Ptg;
import com.valkyrlabs.formats.XLS.formulas.PtgArray;
import com.valkyrlabs.formats.XLS.formulas.PtgExp;
import com.valkyrlabs.formats.XLS.formulas.PtgMemArea;
import com.valkyrlabs.formats.XLS.formulas.PtgRef;
import com.valkyrlabs.toolkit.ByteTools;
import com.valkyrlabs.toolkit.Logger;


/**
 * The Array class describes a formula that was Array-entered into a range of cells.
 * <p>
 * The range of Cells in which the array is entered is defined by the rwFirst, last, colFirst and last fields.
 * <p>
 * The Array record occurs directly after the Formula record for the cell in the upper-left corner of the array -- that is, the cell
 * defined by the rwFirst and colFirst fields.
 * <p>
 * The parsed expression is the array formula -- consisting of Ptgs.
 * <p>
 * You should ignore the chn field when you read the file, it must be 0 if written.
 * <p>
 * <p>
 * OFFSET		NAME			SIZE		CONTENTS
 * -----
 * 4				rwFirst			2			FirstRow of the array
 * 6				fwLast			2			Last Row of the array
 * 8				colFirst		1			First Column of the array
 * 9				colLast			1			Last Column of the array
 * 10			grbit			2			Option Flags
 * 12			chn				4			set to 0, ignore
 * 16			cce				2			Length of the parsed expression
 * 18			rgce			var		Parsed Expression
 * <p>
 * grbit fields
 * bit 1 = fAlwaysCalc - always calc the formula
 * bit 2 = fCalcOnLoad - calc formula on load
 */
public final class Array extends com.valkyrlabs.formats.XLS.XLSRecord {

    private static final long serialVersionUID = -7316545663448065447L;

    private Formula parentRec = null;
    private Stack expression;

    public int getFirstRow() {
        return rwFirst;
    }

    public void setFirstRow(int i) {
        rwFirst = (short) i;
    }

    public int getLastRow() {
        return rwLast;
    }

    public void setLastRow(int i) {
        rwLast = (short) i;
    }

    public int getFirstCol() {
        return colFirst;
    }

    public void setFirstCol(int i) {
        colFirst = (byte) i;
    }

    public int getLastCol() {
        return colLast;
    }

    public void setLastCol(int i) {
        colLast = (byte) i;
    }

    /*
     * For getRow() and getCol() we are going to return the upper-right hand
     * location of the sharedformula.  This should be the same location referenced
     * in the PTGExp associated with these formulas as well
     */
    public int getRowNumber() {
        return getFirstRow();
    }

    public int getCol() {
        return getFirstCol();
    }

    public void init() {
        super.init();
        this.setOpcode(ARRAY);
        rwFirst = (short) com.valkyrlabs.toolkit.ByteTools.readInt(this.getByteAt(0), this.getByteAt(1), (byte) 0, (byte) 0);
        rwLast = (short) com.valkyrlabs.toolkit.ByteTools.readInt(this.getByteAt(2), this.getByteAt(3), (byte) 0, (byte) 0);

        colFirst = (short) ByteTools.readUnsignedShort(this.getByteAt(4), (byte) 0);
        colLast = (short) ByteTools.readUnsignedShort(this.getByteAt(5), (byte) 0);
        grbit = ByteTools.readShort(this.getByteAt(6), this.getByteAt(7));
        cce = ByteTools.readShort(this.getByteAt(12), this.getByteAt(13));
        rgce = this.getBytesAt(14, this.cce);

        expression = ExpressionParser.parseExpression(rgce, this);

        // for PtgArray and PtgMemAreas, have Constant Array Tokens following expression
        // this length is NOT included in cce
        int posExtraData = cce + 14;
        int len = this.getData().length;
        for (int i = 0; i < expression.size(); i++) {
            if (expression.get(i) instanceof PtgArray) {
                try {
                    byte[] b = new byte[(len - posExtraData)];
                    System.arraycopy(this.getData(), posExtraData, b, 0, len - posExtraData);
                    PtgArray pa = (PtgArray) expression.get(i);
                    posExtraData += pa.setArrVals(b);
                } catch (Exception e) {
                    Logger.logErr("Array: error getting Constants " + e.getLocalizedMessage());
                }
            } else if (expression.get(i) instanceof PtgMemArea) {
                try {
                    PtgMemArea pm = (PtgMemArea) expression.get(i);
                    byte[] b = new byte[pm.getnTokens() + 8];
                    System.arraycopy(pm.getRecord(), 0, b, 0, 7);
                    System.arraycopy(this.getData(), posExtraData, b, 7, pm.getnTokens());
                    pm.init(b);
                    posExtraData += pm.getnTokens();
                } catch (Exception e) {
                    Logger.logErr("Array: error getting memarea constants " + e.toString());
                }
            }
        }
        if (DEBUGLEVEL > DEBUG_LOW)
            Logger.logInfo("Array encountered at: " + this.wkbook.getLastbound().getSheetName() + "!" + ExcelTools.getAlphaVal(colFirst) + (rwFirst + 1) + ":" + ExcelTools.getAlphaVal(colLast) + (rwLast + 1));
    }

    /**
     * Associate this record with a worksheet.
     * init array refs as well
     */
    public void setSheet(Sheet b) {
        this.worksheet = b;
        // add to array formula references since this is the parent
        if (expression != null) { // it's been initted
            String loc = ExcelTools.formatLocation(new int[]{rwFirst, colFirst});    // this formula address == array formula references for OOXML usage
            String ref = ExcelTools.formatRangeRowCol(new int[]{rwFirst, colFirst, rwLast, colLast});
            ((Boundsheet) b).addParentArrayRef(loc, ref);        // formula address, array formula references OOXML usage
        }

    }

    /**
     * init Array record from formula string
     *
     * @param fmla
     */
    public void init(String fmla, int rw, int col) {
        // TODO: ever a case of rwFirst!=rwLast, colFirst!=colLast ?????
        this.setOpcode(ARRAY);
        rwFirst = (short) rw;
        rwLast = (short) rw;
        colFirst = (byte) col;
        colLast = (byte) col;
        grbit = 0x2;    // calc on load = default	 20090824 KSC	[BugTracker 2683]
        fmla = fmla.substring(1, fmla.length() - 1);     // parse formula string and add stack to Array record
        Stack newptgs = FormulaParser.getPtgsFromFormulaString(this, fmla);
        expression = newptgs;
        this.updateRecord();

    }

    /*  Update the record byte array with the modified ptg records
     */
    public void updateRecord() {
        // first, get expression length
        cce = 0;        // sum up expression (Ptg) records
        for (int i = 0; i < expression.size(); i++) {
            Ptg ptg = (Ptg) expression.get(i);
            cce += ptg.getRecord().length;
            if (ptg instanceof PtgRef) {
                ((PtgRef) ptg).setArrayTypeRef();
            }
        }

        byte[] newdata = new byte[14 + cce];        // total record data (not including extra data, if any)

        int pos = 0;
        // 20090824 KSC: [BugTracker 2683]
        // use setOpcode rather than setting 1st byte as it's a record not a ptg		newdata[0]= (byte) XLSConstants.ARRAY;			pos++;
        this.setOpcode(ARRAY);
        System.arraycopy(ByteTools.shortToLEBytes(rwFirst), 0, newdata, pos, 2);
        pos += 2;
        System.arraycopy(ByteTools.shortToLEBytes(rwLast), 0, newdata, pos, 2);
        pos += 2;
        newdata[pos++] = (byte) colFirst;
        newdata[pos++] = (byte) colLast;
        System.arraycopy(ByteTools.shortToLEBytes(grbit), 0, newdata, pos, 2);    // 20090824 KSC: Added [BugTracker 2683]

        pos = 12;
        System.arraycopy(ByteTools.shortToLEBytes(cce), 0, newdata, pos, 2);
        pos += 2;

        // expression
        rgce = new byte[cce];                    // expression record data
        pos = 0;
        byte[] arraybytes = new byte[0];
        for (int i = 0; i < expression.size(); i++) {
            Ptg ptg = (Ptg) expression.get(i);
            // trap extra data after expression (not included in cce count) 
            if (ptg instanceof PtgArray) {
                PtgArray pa = (PtgArray) ptg;
                arraybytes = ByteTools.append(pa.getPostRecord(), arraybytes);
            } else if (ptg instanceof PtgMemArea) {
                PtgMemArea pm = (PtgMemArea) ptg;
                arraybytes = ByteTools.append(pm.getPostRecord(), arraybytes);
            }
            byte[] b = ptg.getRecord();
            System.arraycopy(b, 0, rgce, pos, ptg.getLength());
            pos = pos + ptg.getLength();
        }
        System.arraycopy(rgce, 0, newdata, 14, cce);
        newdata = ByteTools.append(arraybytes, newdata);
        this.setData(newdata);
    }

    /**
     * Returns the top left location of the array, used for identifying which array goes with what formula.
     */
    public String getParentLocation() {
        int[] in = new int[2];
        in[0] = colFirst;
        in[1] = rwFirst;
        String s = ExcelTools.formatLocation(in);
        return s;
    }

    /**
     * Determines whether the address is part of the array Formula range
     */
    boolean isInRange(String addr) {
        return com.valkyrlabs.OpenXLS.ExcelTools.isInRange(addr, rwFirst, rwLast, colFirst, colLast);
    }


    public Object getValue(PtgExp pxp) {
        return FormulaCalculator.calculateFormula(expression);
    }

    /**
     * return parent formula
     *
     * @return
     */
    public Formula getParentRec() {
        return parentRec;
    }

    /**
     * link this shared formula to it's parent formula
     */
    public void setParentRec(Formula f) {
        this.parentRec = f;
    }

    /**
     * return the string representation of the array formula
     */
    public String getFormulaString() {
        String expressString = FormulaParser.getExpressionString(expression);
        if (!"".equals(expressString))
            return expressString.substring(1);
        return "";
    }

    /**
     * allow access to expression
     *
     * @return
     */
    public Stack getExpression() {
        return expression;
    }

    /**
     * return the cells referenced by this array in string range form
     *
     * @return
     */
    public String getArrayRefs() {
        int[] rc = new int[2];
        rc[0] = rwFirst;
        rc[1] = colFirst;
        String rowcol1 = ExcelTools.formatLocation(rc);
        rc[0] = rwLast;
        rc[1] = colLast;
        String rowcol2 = ExcelTools.formatLocation(rc);
        return rowcol1 + ":" + rowcol2;
    }
}