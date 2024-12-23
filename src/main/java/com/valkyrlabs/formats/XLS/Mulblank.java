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

import com.valkyrlabs.OpenXLS.CellRange;
import com.valkyrlabs.OpenXLS.ExcelTools;
import com.valkyrlabs.toolkit.ByteTools;
import com.valkyrlabs.toolkit.Logger;

import java.util.*;



/** <b>Mulblank: Multiple Blank Cells (BEh)</b><br>
    This record stores up to 256 BLANK equivalents in
    a space-saving format.

 	TODO: check compatibility with Excel2007 MAXCOLS 

    <p><pre>
    offset  name        size    contents
    ---    
    4       rw          2       Row Number
    6       colFirst    2       Column Number of the first col of multiple Blank record
    8       rgixfe      var     Array of indexes to XF records
    10      colLast     2       Last Column containing Blank objects
    </p></pre>

    * @see Blank
 */

public final class Mulblank extends XLSCellRecord /*implements Mul*/
{
	
    private static final long serialVersionUID = 2707362447402042745L;
	
    short colFirst, colLast; // the colFirst/ColLast indexes determine
    byte[] rgixfe;
    
    public String toString(){
    	return this.getCellAddress();
    }
    
    /** since this is a "MUL" we override this method to
        get a BiffRec Range, not a BiffRec Address.
    */
    public String getCellAddress()
    {
        String retval = "00";
    	if (col==-1) {	// KSC: if not referring to a single cell  
	        int rownum = getRowNumber()+1;
	            retval = ExcelTools.getAlphaVal(colFirst) + String.valueOf(rownum);
	            retval += ":" + ExcelTools.getAlphaVal(colLast) + String.valueOf(rownum);
		} else {	// referring to a single cell
	        int rownum = getRowNumber()+1;
            retval= ExcelTools.getAlphaVal(col) + String.valueOf(rownum);
    	}
        return retval;
    }
    
    /**
     * returns the cell address in int[] {row, col} format
     */
    public int[] getIntLocation() {
    	if (col==-1) {
    		return new int[] {rw, colFirst, rw, colLast };
    	}else{
    		return new int[] {rw, col};
    	}
    }
    
    
    /**
     * return entire range this Mulblank refers to
     * @return
     */
    public String getMulblankRange() {
        String retval = "00";
        int rownum = getRowNumber()+1;
        retval = ExcelTools.getAlphaVal(colFirst) + String.valueOf(rownum);
        retval += ":" + ExcelTools.getAlphaVal(colLast) + String.valueOf(rownum);
        return retval;
    }

        
    /** set the Boundsheet for the Mulblank
        this is needed because Blanks are BiffRec valrecs and
        need to be assigned a BiffRec in the sheet...
            
        the Mulblank itself does not get a cell.
        
    */
    public void setSheet(Sheet bs){
        this.worksheet = bs;
    }
    
	/** initialize this record
	*/
	public void init() {
	    data = getData();
		super.init();
		if (this.getLength() - 4 <= 0) {
			if (DEBUGLEVEL > -1)
				Logger.logInfo("no data in MULBLANK");
		} else {
			rw = ByteTools.readUnsignedShort(this.getByteAt(0), this.getByteAt(1));
			colFirst =
				ByteTools.readShort(this.getByteAt(2), this.getByteAt(3));
			//col = colFirst;
			col= -1;	// flag that this rec hasn't been referred to one cell
			colLast =
				ByteTools.readShort(
					this.getByteAt(this.reclen - 2),
					this.getByteAt(this.reclen - 1));
			//			Sometimes colFirst & colLast are reversed... WTFM$? -jm
			if (colLast < colFirst) {
				short csav = colLast;
				colLast = colFirst;
				colFirst = csav;
				colLast--;
			}
			if (DEBUGLEVEL > 5)
				Logger.logInfo(
					"INFO: MULBLANK range: " + colFirst + ":" + colLast);
			int numblanks = colLast - colFirst + 1;
//			blanks = new ArrayList();
			if (numblanks < 1) {
				Logger.logWarn(
					"WARNING: could not parse Mulblank record: numblanks reported  as:"
						+ numblanks);
				//Logger.logInfo((numblanks >> 12)*-1); ha!
				return;
			}
			rgixfe = this.getBytesAt(4, numblanks * 2);
		}
		// KSC: to use as a blank: 
		this.setIsValueForCell(true);	
		this.isBlank = true;
	}

	/**
	 * returnt the "current" column indicator, if set
	 */
    public short getColNumber(){
    	if (col!=-1) 
    		return col;
    	return colFirst;
    }
    
    /**
     * reset the "current" column use to reference a single blank of this Mulblank range of blank cells
     * @return
     */
    public void resetCol() {
    	col= colFirst;
    }
    /**
     * sets the first column of the range of blank cells referenced by this Mulblank
     * @param c
     */
    public void setColFirst(int c) {
    	colFirst= (short) c;
    }
    
    /**
     * sets the last column of the range of blank cells referenced by this Mulblank
     * @param c
     */
    public void setColLast(int c) {
    	colLast= (short) c;
    }
    
    /**
     * return sthe first column of the range of blank cells referenced by this Mulblank
     * @return
     */
    public int getColFirst() { return colFirst; }
    
    /**
     * return sthe last column of the range of blank cells referenced by this Mulblank
     * @return
     */
    public int getColLast() { return colLast; }
	/**
	 * revise range of cells this Mulblank refers to; return true if no more blanks in range
	 * @param c		col number to remove, 0-based
	 */
    public boolean removeCell(short c) {
		if (c==colFirst) {
			colFirst++;
			byte[] tmp= new byte[rgixfe.length-2];
			System.arraycopy(rgixfe, 2, tmp, 0, tmp.length);	// skip first
			rgixfe= tmp;
		}
		else if (c==colLast) {
			colLast--;
			byte[] tmp= new byte[rgixfe.length-2];
			System.arraycopy(rgixfe, 0, tmp, 0, tmp.length);	// skip last 
			rgixfe= tmp;
		}
		if (c > colFirst && c < colLast) {
			// must break apart Mulblank as now is non-contiguous ...
			// keep first colFirst->c as a MulBlank
			try {				
				// create the blank records
				for (int i = c+1; i <= colLast; i++) {
					byte[] newblank = {0,0,0,0,0,0};	
					// set the row...
					System.arraycopy(this.getBytesAt(0, 2), 0, newblank, 0, 2);
					// set the col...
					System.arraycopy(ByteTools.shortToLEBytes((short) i),0,newblank,2,2);				
					// set the ixfe
					System.arraycopy(rgixfe, ((i-colFirst) * 2), newblank, 4, 2);
					Blank b = new Blank(newblank);
					b.streamer = this.streamer;
					b.setWorkBook(this.getWorkBook());
					b.setSheet(this.getSheet());
					b.setMergeRange(this.getMergeRange(i-colFirst));
					this.getRow().removeCell((short)i);// remove this mulblank from the cells array
					this.getWorkBook().addRecord(b, true);	// and add a blank in it's place
				}
				// truncate the rgixfe:
				byte[] tmp= new byte[(2*(c-colFirst+1))];
				System.arraycopy(rgixfe, 0, tmp, 0, tmp.length);	// skip last 
				rgixfe= tmp;
				// now truncate the Mulblank
				colLast= (short)(c-1);	
			} catch (Exception e) {
				Logger.logInfo("initializing Mulblank failed: " + e);
			}
			col= c;	// the blank to remove
		}
		if (colFirst < 0 || colLast < 0) {	// can happen when removing multiple cells ..?
			return true;
		}
		if (colFirst==colLast) {// covert to a single blank
			byte[] newblank = {0,0,0,0,0,0};	
			// set the row...
			System.arraycopy(this.getBytesAt(0, 2), 0, newblank, 0, 2);
			// set the col...
			System.arraycopy(ByteTools.shortToLEBytes((short) colFirst),0,newblank,2,2);				
			// set the ixfe
			System.arraycopy(rgixfe, 0, newblank, 4, 2);
			Blank b = new Blank(newblank);
			b.streamer = this.streamer;
			b.setWorkBook(this.getWorkBook());
			b.setSheet(this.getSheet());
			b.setMergeRange(this.getMergeRange(colFirst));
			col= colFirst;
			this.getRow().removeCell(this);// remove this mulblank from the cells array
			this.getWorkBook().addRecord(b, true);
			col= c;	// still have to remove cell at col c
			return false;	// no more mulblanks
		}
		updateRecord();
		if (colFirst>colLast)	// no more blanks in range ... can happen??
			return true;	// can delete it
		return false;	// don't delete this rec
	}
	
	
		
    /** return a blank string val 
     */
    public String getStringVal() {return "";}
    
	/**
	 * used to set the cell which this will be referred to, used when trying to access
	 * ixfe 
	 * @param c
	 */
	public void setCurrentCell(short c) {
		col= c;		
	}	    
	/**
	 * NOTE: Mublanks can have a portion of it's blank range which is merged: must determine if 
	 * the current cell is truly part of the merge range ...
	 * @return
	 */
	public CellRange getMergeRange() {
		if (mergeRange==null) return null;
		if (col==-1) 
			return mergeRange;	// this shouldn't happen ...
		if (mergeRange.contains(new int[] {getRowNumber(), col, getRowNumber(), col}))
				return mergeRange;
		return null;	// desired cell is NOT contained within master merge range
	}
	
	/**
	 * retrieves the merged range for the desired cell in this group of blanks
	 * @param col
	 * @return
	 */
	private CellRange getMergeRange(int col) {
		if (mergeRange==null) return null;
		if (col==-1) 
			return mergeRange;	// this shouldn't happen ...
		if (mergeRange.contains(new int[] {getRowNumber(), col, getRowNumber(), col}))
				return mergeRange;
		return null;	// desired cell is NOT contained within master merge range		
	}
	
    /** get the ixfe for the desired referred-to blank 
    */
    public int getIxfe(){
    	int idx= 0;
    	if (col!=-1 && col >= colFirst && col <= colLast) 
			idx= (col-colFirst)*2;    	
		ixfe= ByteTools.readShort(rgixfe[idx], rgixfe[idx+1]);
		myxf= this.getWorkBook().getXf(ixfe);	// set myxf to correct xf for cell in group of mulblanks
        return this.ixfe;
    }
    
    /**
     * Get the referenced columns this mulblank has,
     */
    public ArrayList<Integer> getColReferences(){
    	ArrayList<Integer> colRefs = new ArrayList<Integer>();
    	for (int i=this.colFirst;i<=this.colLast;i++){
    			colRefs.add(i);
    	}
    	return colRefs;
    }

    /**
     * sets the ixfe for the specific cell of the Mulblank
     * (each cell in a series of multiple blanks has their own ixfe)
     */
    public void setIxfe(int i){	
    	int idx= 0;
    	if (col!=-1 && col >= colFirst && col <= colLast) 
			idx= (col-colFirst)*2;
    	
    	byte[] b= ByteTools.shortToLEBytes((short)i);
    	rgixfe[idx]= b[0];
    	rgixfe[idx+1]= b[1];
		updateRecord();
		ixfe= i;		
		myxf= this.getWorkBook().getXf(ixfe);
    }
    /** set the column
     */
     public void setCol(short i){
         col = i;
     }
    
    /**
     * given new info (colFirst, colLast and rgixfe) update data record
     */
    private void updateRecord() {
        byte[] data= new byte[2+2+2 + rgixfe.length];
        data[0]= this.getData()[0];		// row shouldn't have changed
        data[1]= this.getData()[1];
        byte[] b= ByteTools.shortToLEBytes(colFirst);
        data[2]= b[0];					
        data[3]= b[1];
        // after colfirst= rgixfe
        System.arraycopy(rgixfe, 0, data, 4, rgixfe.length);
        b= ByteTools.shortToLEBytes(colLast);
        data[4+rgixfe.length]= b[0];
        data[5+rgixfe.length]= b[1];    	
        setData(data);
    }
    
    public static XLSRecord getPrototype(){
	Mulblank mb= new Mulblank();
	mb.setOpcode(MULBLANK);
	mb.setData(new byte[] { 0, 0, 0, 0, 0, 0});
	mb.col= -1;
	return mb;
    }    
    /** returns the number of ixfe fields
    */
    int getNumFields(){
        return (colLast - colFirst) + 1;
    }
    
    

}