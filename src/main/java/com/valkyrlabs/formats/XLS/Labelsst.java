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

import java.util.Calendar;
import java.util.GregorianCalendar;

import com.valkyrlabs.OpenXLS.DateConverter;
import com.valkyrlabs.OpenXLS.WorkBookHandle;
import com.valkyrlabs.toolkit.ByteTools;
import com.valkyrlabs.toolkit.Logger;


/** <b>Labelsst: BiffRec Value, String Constant/Sst 0xFD</b><br>
    The Labelsst record contains a string constant
    from the Shared String Table (Sst).
    The isst field contains a zero-based index into the shared string table
    
    <pre>
    offset  name        size    contents
    ---    
    4       rw          2       Row Number
    6       col         2       Column Number
    8       ixfe        2       Index to XF format record
    10      isst        4       Index into the Sst record
    </pre>
    
    * @see Sst
    * @see Labelsst
    * @see Extsst
*/
public final class Labelsst
extends XLSCellRecord {
    /** 
	* serialVersionUID
	*/
	private static final long serialVersionUID = 467127849827595055L;
	int isst;

    void setIsst(int i){
        isst = i;
        System.arraycopy(ByteTools.cLongToLEBytes(isst), 0, this.getData(), 6, 4);   
        try {
            this.getWorkBook().getSharedStringTable().initSharingOnStrings(isst);
        }catch(NullPointerException e) { ; }
    }
    
    /** Constructor which takes a number value
        an Sst to store its Unicodestring in,
        and returns an int offset to the string
        in the Sst.
    */
    public static Labelsst getPrototype(String val, WorkBook bk) {
        Labelsst retlab = new Labelsst();
        // associate with the Sst
        retlab.originalsize = 10;
        retlab.setOpcode(LABELSST);
        retlab.setLength((short) 10);
		retlab.setData(new byte[retlab.originalsize]);
        //retlab.setDataContainsHeader(true);
		if (val!=null) { // for XLSX handling ... label is linked to sst later
	        // get the high Sst index, insert the new Unicodestring
	    	Sst sst= bk.getSharedStringTable();        
        	retlab.isst  = sst.insertUnicodestring(val);
            System.arraycopy(ByteTools.cLongToLEBytes(retlab.isst), 0, retlab.getData(), 6, 4);
		} else 
			retlab.isst= -1;	// flag it's not set - MUST be set later
        retlab.getData()[4] = 0x0f;
        retlab.setWorkBook(bk);
        retlab.init();
        return retlab;
    }

	public void init() {
		super.init();
		
		// this.initCacheBytes(0,10);
		// get the row, col and ixfe information
		super.initRowCol();
		short s = ByteTools.readShort(this.getByteAt(4), this.getByteAt(5));
		ixfe = (int) s;
		// get the length of the string
		isst =
			ByteTools.readInt(
				this.getByteAt(6),
				this.getByteAt(7),
				this.getByteAt(8),
				this.getByteAt(9));
		this.setIsValueForCell(true);
		this.isString = true;
		this.resetCacheBytes();
		// init shared string info.
		if (isst!=-1){// not initialized - OOXML use - MUST be set later using setIsst
        try {
                this.getWorkBook().getSharedStringTable().initSharingOnStrings(isst);            
        }catch(NullPointerException e) {
            // nothing.  When adding new strings we have access issues, but it doesn't matter, we just care on book initialization for this..
        }
		}
	}

    
	private Unicodestring unsharedstr;
	
	
	void initUnsharedString(){
		unsharedstr = this.getWorkBook().getSharedStringTable().getUStringAt(isst);
	}
	
	public Unicodestring getUnsharedString(){
	    if(unsharedstr ==null)
			this.initUnsharedString();		
		return unsharedstr;
	}
    
    /**
     * Adds the LabelSST's string to the sst.  
     * 
     * This is used when a worksheet is transferred over to a book that
     *  does not contain it's entry in the sst.
     */
    boolean insertUnsharedString(Sst sst) {
        if (unsharedstr == null) {
            return false;
        }
        this.isst = sst.insertUnicodestring(unsharedstr.toString());
        this.setIsst(isst);
        return true;
    }
    
    /** Returns the value of the Unicodestring
        int the Shared String Table pointed to by this
        LABELSst record.
     */
	public String getStringVal(){
	    if(unsharedstr !=null)
	    	return unsharedstr.toString();
	    else 
	    	return this.getWorkBook().getSharedStringTable().getUStringAt(isst).toCachingString();
    }
	
	/**
	 * try to convert the String Value of this Labelsst record to an int
	 * If it cannot be converted, returns NaN.
	 */
	public int getIntVal() {
		String s= getStringVal();
		try {
			Integer i= Integer.valueOf(s);
			return i.intValue();
		} catch (NumberFormatException n) {
			return (int)Float.NaN;
		}
	}
	
	/**
	 * try to convert the String Value of this Labelsst record to a double
	 * If it cannot be converted,return NaN.
	 */
	public double getDblVal() {
		String s= getStringVal();
		try {
			Double d= new Double(s);
			return d.doubleValue();
		} catch (NumberFormatException n) {
			this.getXfRec();
			if (myxf.isDatePattern()) { // use it
				try {
					String format= myxf.getFormatPattern();
					WorkBookHandle.simpledateformat.applyPattern(format);
					java.util.Date d= WorkBookHandle.simpledateformat.parse(s);
					Calendar c= new GregorianCalendar();
					c.setTime(d);
					if (c==null) return Double.NaN;
					return DateConverter.getXLSDateVal(c);
				} catch (Exception e) {
					// fall through
				}
			} 
            Calendar c= DateConverter.convertStringToCalendar(s);
            if (c==null) return Double.NaN;
            return DateConverter.getXLSDateVal(c);
		}
	}

    /** set a new value for the string
    */
	public void setStringVal(String v){
		String ov = this.getStringVal();
        if(v.equals(ov))return;
        if(this.getSheet().getWorkBook().getSharedStringTable().isSharedString(isst)) {
            isst  = this.getSheet().getWorkBook().getSharedStringTable().insertUnicodestring(v);
            System.arraycopy(ByteTools.cLongToLEBytes(isst), 0, getData(), 6, 4);
            init();
            // reset unsharedstr (see getStringVal) specifically to fix OOXML t="s" setStringVal 
            Unicodestring str = this.getSheet().getWorkBook().getSharedStringTable().getUStringAt(isst);
            this.unsharedstr= str; 
        }else {
            Unicodestring str = this.getSheet().getWorkBook().getSharedStringTable().getUStringAt(isst);
            //ensure reclen and datalen are maintained correctly:
            int origLen= str.getLength();
            str.updateUnicodeString(v);
            int delta= str.getLength()-origLen;
            this.getSheet().getWorkBook().getSharedStringTable().adjustSstLength(delta);
            this.unsharedstr = str;
        }
    }
	
		/** set this Label cell to a new Unicode string
		 *  Rich Unicode strings include formatting information  
	    */
		// 20090520 KSC: for OOXML, must use entire Unicode string so retain formatting info
		public void setStringVal(Unicodestring v){
			if (v.equals(this.getUnsharedString()))
				return;
			isst= this.getSheet().getWorkBook().getSharedStringTable().find(v);	// find this particular unicode string (including formatting) 
			if (isst==-1)
				isst= this.getSheet().getWorkBook().getSharedStringTable().insertUnicodestring(v);
            System.arraycopy(ByteTools.cLongToLEBytes(isst), 0, getData(), 6, 4);
            init();
            this.unsharedstr= v; 
	    }
	

	/**
	 * return string representation
	 */
	public String toString() {
		try{
			return "LABELSST:" + this.getCellAddress() + ":" + getStringVal();
		}catch(Exception e){
			Logger.logErr("Labelsst toString failed.",e);
			return "#ERR!";
		}
	} 
}
