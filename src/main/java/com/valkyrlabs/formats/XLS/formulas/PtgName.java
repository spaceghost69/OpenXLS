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

import com.valkyrlabs.OpenXLS.FormulaNotFoundException;
import com.valkyrlabs.formats.XLS.Formula;
import com.valkyrlabs.formats.XLS.Name;
import com.valkyrlabs.formats.XLS.WorkBook;
import com.valkyrlabs.toolkit.ByteTools;
import com.valkyrlabs.toolkit.FastAddVector;
import com.valkyrlabs.toolkit.Logger;


/**
 * This PTG stores an index to a name.  The ilbl field is a 1 based index to the table
 * of NAME records in the workbook
 * <p>
 * OFFSET      NAME        sIZE        CONTENTS
 * ---------------------------------------------
 * 0           ilbl        2           Index to the NAME table
 * 2           (reserved)  2   `       Must be 0;
 *
 * @see Ptg
 * @see Formula
 */
public class PtgName extends GenericPtg implements Ptg, IlblListener {

    /**
     * serialVersionUID
     */
    private static final long serialVersionUID = 8047146848365098162L;
    short ilbl;
    String name;

    //default constructor
    public PtgName() {
        ptgId = 0x23;        // reference type is default
    }

    // 20100218 KSC:
    // constructor which sets a specific id
    // to specify whether this PtgName is of value, ref or array type
    // (PtgNameV, PtgNameR or PtgNameA)
    public PtgName(int id) {
        ptgId = (byte) id;
//				0x23=   Ref
//		ptgId = 0x43;	Value
    }

    public boolean getIsOperand() {

        return true;
    }

    // lookup Name object  in Workbook and return handle
    public Name getName() {
        WorkBook b = this.getParentRec().getWorkBook();
        Name n = null;
        try {
            n = b.getName(ilbl);
        } catch (Exception ex) {
        }
        return n;
    }

    /**
     * For creating a ptg name from formula parser
     */
    public void setName(String name) {
        record = new byte[5];
        record[0] = ptgId;
        WorkBook b = this.getParentRec().getWorkBook();
        ilbl = (short) b.getNameNumber(name);
        this.addListener();
        record[1] = (byte) ilbl;
    }

    public void init(byte[] b) {
        ptgId = b[0];
        record = b;
        this.populateVals();
        addToRefTracker();
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
                ptgId = 0x43;
                break;
            case REFERENCE:
                ptgId = 0x23;
                break;
            case Ptg.ARRAY:
                ptgId = 0x63;
                break;
        }
        record[0] = ptgId;
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
                parent_rec.getWorkBook().getRefTracker().addPtgNameReference(this);
        } catch (Exception ex) {
            Logger.logErr("PtgRef.addToRefTracker() failed.", ex);
        }
    }

    private void populateVals() {
        ilbl = ByteTools.readShort(record[1], record[2]);
    }

    public int getVal() {
        return ilbl;
    }

    public void setVal(int i) {
        ilbl = (short) i;
        this.updateRecord();
    }

    public short getIlbl() {
        return ilbl;
    }

    public void setIlbl(short i) {
        if (ilbl != i) {
            ilbl = i;
            this.updateRecord();
        }
    }

    public void storeName(String nm) {
        name = nm;
    }

    /*
     *
     * returns the string value of the name
		@see com.valkyrlabs.formats.XLS.formulas.Ptg#getValue()
     */
    public Object getValue() {
        Name n = getName();
        try {
            Ptg[] p = n.getCellRangePtgs();
            if (p.length == 0) {
                return "#NAME?";
            } else if (p.length == 1 || !(this.parent_rec instanceof com.valkyrlabs.formats.XLS.Array)) {    // usual case
                return p[0].getValue();
            } else { // multiple values; create an array
                String retarry = "";
                for (int i = 0; i < p.length; i++) {
                    retarry = retarry + p[i].getValue() + ",";
                }
                retarry = "{" + retarry.substring(0, retarry.length() - 1) + "}";
                PtgArray pa = new PtgArray();
                pa.setVal(retarry);
                return pa;
            }
        } catch (Exception e) {
        }
        //    	String s = n.getName();
        //return n;
        return "#NAME?";
    }

    public String getTextString() {
        Name n = getName();
        if (n == null)
            return "#NAME!";
        return n.getName();
    }

    public String getStoredName() {
        return name;
    }

    public void updateRecord() {
        byte[] brow = ByteTools.cLongToLEBytes(ilbl);
        record[1] = brow[0];
        record[2] = brow[1];
        if (parent_rec != null) {
            if (parent_rec instanceof Formula)
                ((Formula) parent_rec).updateRecord();
        }
    }

    /**
     * Override due to mystery extra byte
     * occasionally found in ptgName recs.
     */
    public int getLength() {
        if (record != null) return record.length;
        return PTG_NAME_LENGTH;
    }

    public String toString() {
        if (this.getName() != null)
            return this.getName().getName();
        return "[Null]";
    }

    public Ptg[] getComponents() {
        FastAddVector v = new FastAddVector();
        Ptg p = this.getName().getPtga();
        Ptg[] pcomps = p.getComponents();
        if (pcomps != null) {
            for (int x = 0; x < pcomps.length; x++) {
                v.add(pcomps[x]);
            }
        } else {
            v.add(p);
        }
        Ptg[] retPtgs = new Ptg[v.size()];
        retPtgs = (Ptg[]) v.toArray(retPtgs);
        return retPtgs;
    	
/*    	Ptg[] p = this.getName().getComponents();
    	for (int i=0;i<p.length;i++){
    	    Ptg[] pcomps = p[i].getComponents();
    		if (pcomps!= null){
    			for (int x=0;x<pcomps.length;x++){
    				v.add(pcomps[x]);    				
    			}
    		}else{
    			v.add(p[i]);
    		}
    	}
    	Ptg[] retPtgs = new Ptg[v.size()];
    	retPtgs = (Ptg[])v.toArray(retPtgs);
    	return retPtgs;*/
    }

    /**
     * return referenced Names' location
     *
     * @see com.valkyrlabs.formats.XLS.formulas.GenericPtg#getLocation()
     */
    public String getLocation() throws FormulaNotFoundException {
        if (this.getName() != null)
            try {
                return this.getName().getLocation();
            } catch (Exception e) {
            }
        return null;
    }

    public void addListener() {
        Name n = this.getName();
        if (n != null) {
            n.addIlblListener(this);
            this.storeName(n.getName());
        }

    }
}
    
    