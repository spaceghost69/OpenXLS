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


import com.valkyrlabs.OpenXLS.DateConverter;
import com.valkyrlabs.OpenXLS.FormulaNotFoundException;
import com.valkyrlabs.OpenXLS.FunctionNotSupportedException;
import com.valkyrlabs.formats.XLS.*;
import com.valkyrlabs.toolkit.ByteTools;
import com.valkyrlabs.toolkit.Logger;

import java.util.Calendar;

public abstract class GenericPtg
        implements Ptg, Cloneable {

    public static final long serialVersionUID = 666555444333222l;
    // Parent Rec is the BiffRec record referenced by Operand Ptgs
    protected XLSRecord parent_rec;
    double doublePrecision = 0.00000001;        // doubles/floats cannot be compared for exactness so use precision comparator
    byte ptgId;
    byte[] record;
    Ptg[] vars = null;
    int lock_id = -1;
    private int locationLocked = Ptg.PTG_LOCATION_POLICY_UNLOCKED;
    private BiffRec trackercell = null;

    /**
     * Returns an array of doubles from number-type ptg's sent in.
     * This should only be referenced by sub-classes.
     * <p>
     * Null values accessed are treated as 0.  Within excel (empty cell values == 0) Tested!
     * Sometimes as well you can get empty string values, "".  These are NOT EQUAL ("" != 0)
     *
     * @param pthings
     * @return
     */
    protected static Object[] getValuesFromPtgs(Ptg[] pthings) {
        Object[] obar = new Object[pthings.length];
        for (int t = 0; t < obar.length; t++) {
            if (pthings[t] instanceof PtgErr)
                return null;
            if (pthings[t] instanceof PtgArray) {
                obar[t] = pthings[t].getComponents();    // get all items in array as Ptgs
                Object v = null;
                try {
                    v = getValuesFromObjects((Object[]) obar[t]);    // get value array from the ptgs
                } catch (NumberFormatException e) {    // string or non-numeric values
                    v = getStringValuesFromPtgs((Ptg[]) obar[t]);
                }
                obar[t] = v;
            } else {
                Object pval = pthings[t].getValue();
                if (pval instanceof PtgArray) {
                    obar[t] = ((PtgArray) pval).getComponents();    // get all items in array as Ptgs
                    Object v = null;
                    try {
                        v = getValuesFromObjects((Object[]) obar[t]);    // get value array from the ptgs
                    } catch (NumberFormatException e) {    // string or non-numeric values
                        v = getStringValuesFromPtgs((Ptg[]) obar[t]);
                    }
                    obar[t] = v;
                } else if (pval instanceof Name) {    // then get it's components ...
                    obar[t] = pthings[t].getComponents();
                    Object v = null;
                    try {
                        v = getValuesFromPtgs((Ptg[]) obar[t]);    // get value array from the ptgs
                    } catch (NumberFormatException e) {    // string or non-numeric values
                        v = getStringValuesFromPtgs((Ptg[]) obar[t]);
                    }
                    obar[t] = v;
                } else {    // it's a single value
                    try {
                        obar[t] = new Double(getDoubleValueFromObject(pval));
                    } catch (NumberFormatException e) {
                        if (pval instanceof CalculationException)
                            obar[t] = pval.toString();
                        else
                            obar[t] = pval;
                    }
                }
            }
        }
        return obar;
    }

    /**
     * Returns an array of doubles from number-type ptg's sent in.
     * This should only be referenced by sub-classes.
     * <p>
     * Null values accessed are treated as 0.  Within excel (empty cell values == 0) Tested!
     * Sometimes as well you can get empty string values, "".  These are NOT EQUAL ("" != 0)
     *
     * @param pthings
     * @return
     */
    protected static double[] getValuesFromObjects(Object[] pthings) throws NumberFormatException {
        double[] returnDbl = new double[pthings.length];
        for (int i = 0; i < pthings.length; i++) {

            // Object o = pthings[i].getValue();
            Object o = pthings[i];

            if (o == null) {    // NO!! "" is NOT "0", blank is, but not a zero length string.  Causes calc errors, need to handle diff somehow20081103 KSC: don't error out if "" */
                returnDbl[i] = 0.0;
            } else if (o instanceof Double) {
                returnDbl[i] = ((Double) o).doubleValue();
            } else if (o instanceof Integer) {
                returnDbl[i] = ((Integer) o).intValue();
            } else if (o instanceof Boolean) {    // Excel converts booleans to numbers in calculations 20090129 KSC
                returnDbl[i] = (((Boolean) o).booleanValue() ? 1.0 : 0.0);
            } else if (o instanceof PtgBool) {
                returnDbl[i] = (((Boolean) (((PtgBool) o).getValue())).booleanValue() ? 1.0 : 0.0);
            } else if (o instanceof PtgErr) {
                // ?
            } else {
                String s = o.toString();
                Double d = new Double(s);
                returnDbl[i] = d.doubleValue();
            }
        }
        return returnDbl;
    }

    /**
     * convert a value to a double, throws exception if cannot
     *
     * @param o
     * @return double value if possible
     * @throws NumberFormatException
     */
    public static double getDoubleValue(Object o, XLSRecord parent)
            throws NumberFormatException {
        if (o instanceof Double)
            return ((Double) o).doubleValue();
        if (o == null || o.toString().equals("")) {
            // empty string is interpreted as 0 if show zero values
            if (parent != null && parent.getSheet().getWindow2().getShowZeroValues())
                return 0.0;
            // otherwise, throw error
            throw new NumberFormatException();
        }
        return new Double(o.toString()).doubleValue();    // will throw NumberFormatException if cannot convert
    }

    /**
     * converts a single Ptg number-type value to a double
     */
    public static double getDoubleValueFromObject(Object o) {
        double ret = 0.0;
        if (o == null) {    // 20081103 KSC: don't error out if "" */
            ret = 0.0;
        } else if (o instanceof Double) {
            ret = ((Double) o).doubleValue();
        } else if (o instanceof Integer) {
            ret = ((Integer) o).intValue();
        } else if (o instanceof Boolean) {    // Excel converts booleans to numbers in calculations 20090129 KSC
            ret = (((Boolean) o).booleanValue() ? 1.0 : 0.0);
        } else if (o instanceof PtgErr) {
            // ?
        } else {
            String s = o.toString();
            // handle formatted dates from fields like TEXT() calcs
            if (s.indexOf("/") > -1) {
                try {
                    Calendar c = DateConverter.convertStringToCalendar(s);
                    if (c != null) ret = DateConverter.getXLSDateVal(c);
                } catch (Exception e) {//guess not
                }
            }
            if (ret == 0.0) {
                Double d = new Double(s);
                ret = d.doubleValue();
            }
        }
        return ret;
    }

    /**
     * returns an array of strings from ptg's sent in.
     * This should only be referenced by sub-classes.
     */
    protected static String[] getStringValuesFromPtgs(Ptg[] pthings) {
        String[] returnStr = new String[pthings.length];
        for (int i = 0; i < pthings.length; i++) {
            if (pthings[i] instanceof PtgErr)
                return new String[]{"#VALUE!"};    // 20081202 KSC: return error value ala Excel

            Object o = pthings[i].getValue();
            if (o != null) { // 20070215 KSC: avoid nullpointererror
                try {    // 20090205 KSC: try to convert numbers to ints when converting to string as otherwise all numbers come out as x.0
                    returnStr[i] = String.valueOf(((Double) o).intValue());
                } catch (Exception e) {
                    String s = o.toString();
                    returnStr[i] = s;
                }
            } else
                returnStr[i] = "null"; // 20070216 KSC: Shouldn't match empty string!
        }
        return returnStr;
    }

    /**
     * return properly quoted sheetname
     *
     * @param s
     * @return
     */
    public static final String qualifySheetname(String s) {
        if (s == null || s.equals("")) return s;
        try {
            if (s.charAt(0) != '\'' && (s.indexOf(' ') > -1 || s.indexOf('&') > -1 || s.indexOf(',') > -1 || s.indexOf('(') > -1)) {
                if (s.indexOf("'") == -1)    // normal case of no embedded ' s
                    return "'" + s + "'";
                return "\"" + s + "\"";
            }
        } catch (StringIndexOutOfBoundsException e) {
        }
        return s;
    }

    /**
     * return cell address with $'s e.g.
     * cell AB12 ==> $AB$12
     * cell Sheet1!C2=>Sheet1!$C$2
     * Does NOT handle ranges
     *
     * @param s
     * @return
     */
    public static String qualifyCellAddress(String s) {
        String prefix = "";
        if (s.indexOf("$") == -1) {    // it's not qualified yet
            int i = s.indexOf("!");
            if (i > -1) {
                prefix = s.substring(0, i + 1);
                s = s.substring(i + 1);
            }
            s = "$" + s;
            i = 1;
            while (i < s.length() && !Character.isDigit(s.charAt(i++))) ;
            i--;
            if (i > 0 && i < s.length())
                s = s.substring(0, i) + "$" + s.substring(i);
        }
        return prefix + s;
    }

    public static int getArrayLen(Object o) {
        int len = 0;
        if (o instanceof double[])
            len = ((double[]) o).length;
        return len;
    }

    public Object clone() {
        try {
            return super.clone();
        } catch (CloneNotSupportedException e) {
            // This is, in theory, impossible
            return null;
        }
    }

    /**
     * a locking mechanism so that Ptgs are not endlessly
     * re-calculated
     *
     * @return
     */
    public int getLock() {
        return lock_id;
    }

    /**
     * a locking mechanism so that Ptgs are not endlessly
     * re-calculated
     *
     * @return
     */
    public void setLock(int x) {
        lock_id = x;
    }

    // determine behavior
    public boolean getIsOperator() {
        return false;
    }

    public boolean getIsBinaryOperator() {
        return false;
    }

    public boolean getIsUnaryOperator() {
        return false;
    }

    public boolean getIsStandAloneOperator() {
        return false;
    }

    public boolean getIsPrimitiveOperator() {
        return false;
    }
        
/* ################################################### EXPLANATION ###################################################
   
    1. set string varetvar in all Ptgs
    2. varetvar goes between ptg return vals if any
    3. if this is a funcvar then we loop ptgs and out 
    4. when we call getString or evaluate, we loop into the
        recursive tree and execute on up.
  
   ################################################### EXPLANATION ###################################################*/

    public boolean getIsOperand() {
        return false;
    }

    public boolean getIsFunction() {
        return false;
    }

    public boolean getIsControl() {
        return false;
    }

    public boolean getIsArray() {
        return false;
    }

    public boolean getIsReference() {
        return false;
    }

    /**
     * returns the Location Policy of the Ptg is locked
     * used during automated BiffRec movement updates
     *
     * @return int
     */
    public int getLocationPolicy() {
        return locationLocked;
    }

    /**
     * lock the Location of the Ptg so that it will not
     * be updated during automated BiffRec movement updates
     *
     * @param b setting of the lock the location policy for this Ptg
     */
    public void setLocationPolicy(int b) {
        locationLocked = b;
    }

    /**
     * update the Ptg
     */
    public void updateRecord() {

    }

    /**
     * Returns the number of Params to pass to the Ptg
     */
    public int getNumParams() {
        if (getIsPrimitiveOperator()) return 2;
        return 0;
    }

    /**
     * Operator Ptgs take other Ptgs as arguments
     * so we need to pass them in to get a meaningful
     * value.
     */
    public void setVars(Ptg[] parr) {
        this.vars = parr;
    }

    /*
        Return all of the cells in this range as an array
        of Ptg's.  This is used for range calculations.
    */
    public Ptg[] getComponents() {
        return null;
    }

    /**
     * pass  in arbitrary number of values (probably other Ptgs)
     * and return the resultant value.
     * <p>
     * This effectively calculates the Expression.
     */
    public Object evaluate(Object[] obj) {
        // do something useful
        return this.getString();
    }

    /**
     * return the human-readable String representation of
     * this ptg -- if applicable
     */
    public String getTextString() {

        String strx = "";

        try {
            strx = getString();
        } catch (Exception e) {
            Logger.logErr("Function not supported: " + this.parent_rec.toString());
        }

        if (strx == null)
            return "";

        StringBuffer out = new StringBuffer(strx);
        if (vars != null) {
            int numvars = vars.length;
            if (this.getIsPrimitiveOperator() && this.getIsUnaryOperator()) {
                if (numvars > 0)
                    out.append(vars[0].getTextString());

            } else if (this.getIsPrimitiveOperator()) {
                out.setLength(0);
                for (int x = 0; x < numvars; x++) {
                    out.append(vars[x].getTextString());
                    if (x + 1 < numvars) out.append(this.getString());
                }
            } else if (this.getIsControl()) {
                for (int x = 0; x < numvars; x++) {
                    out.append(vars[x].getTextString());
                }
            } else {
                for (int x = 0; x < vars.length; x++) {
                    if (!(x == 0 && vars[x] instanceof PtgNameX)) {    // KSC: added to skip External name reference for Add-in Formulas
                        String part = vars[x].getTextString();
                        // 20060408 KSC: added quoting in PtgStr.getTextString
//	                    if (vars[x] instanceof PtgStr) // 20060214 KSC: Quote string params
//	                    	part= "\"" + part + "\"";
                        out.append(part);
                        /*if(!part.equals(""))*/
                        out.append(",");
                    }
                }
                if (vars.length > 0) // don't strip 1st paren if no params!  20060501 KSC
                    out.setLength(out.length() - 1); // strip trailing comma
            }
        }
        out.append(getString2());
        return out.toString();
    }

    /*text1 and 2 for this Ptg
     */
    public String getString() {
        return toString();
    }

    /**
     * return the human-readable String representation of
     * the "closing" portion of this Ptg
     * such as a closing parenthesis.
     */

    public String getString2() {
        if (this.getIsPrimitiveOperator()) return "";
        if (this.getIsOperator()) return ")";
        return "";
    }

    public byte getOpcode() {
        return ptgId;
    }

    public void init(byte[] b) {
        ptgId = b[0];
        record = b;
    }

    /**
     * return a Ptg  consisting of the calculated values
     * of the ptg's passed in.  Returns null for any non-operand
     * ptg.
     *
     * @throws CalculationException
     */
    public Ptg calculatePtg(Ptg[] parsething) throws FunctionNotSupportedException, CalculationException {
        return null;

    }

    /**
     * Gets the (return) value of this Ptg as an operand Ptg.
     */
    public Ptg getPtgVal() {
        Object value = this.getValue();
        if (value instanceof Ptg) return (Ptg) value;
        else if (value instanceof Boolean)
            return new PtgBool(((Boolean) value).booleanValue());
        else if (value instanceof Integer)
            return new PtgInt(((Integer) value).intValue());
        else if (value instanceof Number)
            return new PtgNumber(((Number) value).doubleValue());
        else if (value instanceof String)
            return new PtgStr((String) value);
        else return new PtgErr(PtgErr.ERROR_VALUE);
    }

    /**
     * returns the value of an operand ptg.
     *
     * @return null for non-operand Ptg.
     */
    public Object getValue() {
        return null;
    }

    /**
     * Gets the value of the ptg represented as an int.
     * <p>
     * This can result in loss of precision for floating point values.
     * <p>
     * overridden in PtgInt to natively return value.
     *
     * @return integer representing the ptg, or NAN
     */
    public int getIntVal() {
        try {
            return new Double(this.getValue().toString()).intValue();
        } catch (NumberFormatException e) {
            // we should be throwing something better
            if (!(this instanceof PtgErr))    // don't report an error if it's already an error
                Logger.logErr("GetIntVal failed for formula: " + this.getParentRec().toString() + " " + e);
            return 0;
            ///  RIIIIGHT!  throw new FormulaCalculationException();
        }
    }

    /**
     * Gets the value of the ptg represented as an double.
     * <p>
     * This can result in loss of precision for floating point values.
     * <p>
     * NAN will be returned for values that are not translateable to an double
     * <p>
     * overrideen in PtgNumber
     *
     * @return integer representing the ptg, or NAN
     */
    public double getDoubleVal() {
        Object pob = null;
        Double d = null;
        try {
            pob = this.getValue();
            if (pob == null) {
                Logger.logErr("Unable to calculate Formula at " + this.getLocation());
                return java.lang.Double.NaN;
            }
            d = (Double) pob;
        } catch (ClassCastException e) {
            try {
                Float f = (Float) pob;
                d = new Double(f.doubleValue());
            } catch (ClassCastException e2) {
                try {
                    Integer in = (Integer) pob;
                    d = new Double(in.doubleValue());
                } catch (Exception e3) {
                    if (pob == null || pob.toString().equals("")) {
                        d = new Double(0);
                    } else {
                        try {
                            Double dd = new Double(pob.toString());
                            return dd.doubleValue();
                        } catch (Exception e4) {// Logger.logWarn("Error in Ptg Calculator getting Double Value: " + e3);
                            return java.lang.Double.NaN;
                        }
                    }
                }
            }
        } catch (Throwable exp) {
            Logger.logErr("Unexpected Exception in PtgCalculator.getDoubleValue()", exp);
        }
        return d.doubleValue();
    }

    /**
     * So, here you see we can get the static type from the record itself
     * then format the output record.  Some shorthand techniques are shown.
     */
    public byte[] getRecord() {
        return record;
    }

    public String getLocation() throws FormulaNotFoundException {
        return null;
    }

    // these do nothing here...
    public void setLocation(String s) {
    }

    public int[] getIntLocation() throws FormulaNotFoundException {
        return null;
    }

    public XLSRecord getParentRec() {
        return parent_rec;
    }

    public void setParentRec(XLSRecord f) {
        parent_rec = f;
    }

    /**
     * if the Ptg needs to keep a handle to a cell, this is it...
     * tells the Ptg to get it on its own...
     */
    public void updateAddressFromTrackerCell() {
        this.initTrackerCell();
        BiffRec trk = getTrackercell();
        if (trk != null) {
            String nad = trk.getCellAddress();
            setLocation(nad);
        }
    }

    /**
     * if the Ptg needs to keep a handle to a cell, this is it...
     * tells the Ptg to get it on its own...
     */
    public void initTrackerCell() {
        if (getTrackercell() == null) {
            try {
                BiffRec trk = this.getParentRec().getSheet().getCell(this.getLocation());
                setTrackercell(trk);
            } catch (Exception e) {
                Logger.logWarn("Formula reference could not initialize:" + e.toString());
            }
        }
    }

    /**
     * @return Returns the trackercell.
     */
    public BiffRec getTrackercell() {
        return trackercell;
    }

    /**
     * @param trackercell The trackercell to set.
     */
    public void setTrackercell(BiffRec trackercell) {
        this.trackercell = trackercell;
    }

    //TODO: PtgRef.isBlank should override!
    public boolean isBlank() {
        return false;
    }

    /**
     * generic reading of a row byte pair with handling for Excel 2007 if necessary
     *
     * @param b0
     * @param b1
     * @return int row
     */
    public int readRow(byte b0, byte b1) {
        if ((parent_rec != null && !parent_rec.getWorkBook().getIsExcel2007())) {
            int rw = com.valkyrlabs.toolkit.ByteTools.readInt(b0, b1, (byte) 0, (byte) 0);
            if (rw >= MAXROWS_BIFF8 - 1 || rw < 0 || this instanceof PtgRefN)    // PtgRefN's are ALWAYS relative and therefore never over 32xxx
                rw = ByteTools.readShort(b0, b1);
            return rw;
        }
        // issue when reading Excel2007 rw from bytes as limits exceed ... try to interpret as best one can
        int rw = com.valkyrlabs.toolkit.ByteTools.readInt(b0, b1, (byte) 0, (byte) 0);
        if (rw == 65535) {    // have to assume that this means a wholeCol reference
            rw = -1;
            ((PtgRef) this).wholeCol = true;
        }
        return rw;
    }

    /**
     * clear out object references in prep for closing workbook
     */
    public void close() {
        parent_rec = null;
        trackercell = null;
        // vars??

    }


} 