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

import java.util.ArrayList;

import com.valkyrlabs.OpenXLS.CellRange;
import com.valkyrlabs.OpenXLS.ExcelTools;
import com.valkyrlabs.formats.XLS.SxAddl.SxcCache;
import com.valkyrlabs.toolkit.ByteTools;
import com.valkyrlabs.toolkit.Logger;

/**
 * SXStreamID 0xD5
 * The SXStreamID record specifies the start of the stream in the PivotCache storage.
 *
   *idStm (2 bytes): An unsigned integer that specifies a stream in the PivotCache storage. The stream specified is the one that has its name equal to the hexadecimal representation of this field. The four-digit hexadecimal string representation of this field, where each hexadecimal letter digit is a capital letter, MUST be equal to the name of a stream (1) in the PivotCache storage.

 */
public class SxStreamID extends XLSRecord implements XLSConstants {
    /** 
	* serialVersionUID
	*/
	private short streamId=-1;
	private static final long serialVersionUID = 2639291289806138985L;
	private ArrayList subRecs= new ArrayList();
	/**
	 * init method
	 */
	public void init(){
        super.init();        
        streamId= ByteTools.readShort(this.getByteAt(0),this.getByteAt(1));
        if (DEBUGLEVEL > 3) Logger.logInfo("SXSTREAMID: streamid:" + streamId);
	}
	
	/**
	 * store like cache-related records under SxStreamID
	 * @param r
	 */
    public void addSubrecord(BiffRec r) {
    	subRecs.add(r);
    }
	
	/**
	 * creates new, default SxStreamID
	 * @return
	 */
    public static XLSRecord getPrototype() {
    	SxStreamID ss = new SxStreamID();
    	ss.setOpcode(SXSTREAMID);
    	ss.setData(new byte[] {0, 0});
    	ss.init();
    	return ss;
    }
    /**
     * returns the streamId -- index linked to appropriate SxView pivot table view
     * @return
     */
	public short getStreamID() { return streamId; }
	
	/**
	 * sets the streamId -- index linked to approriate SxView pivot table view
	 * @param sid
	 */
	public void setStreamID(int sid) {
		streamId= (short) sid;
		byte[] b= ByteTools.shortToLEBytes(streamId);
		this.getData()[0]= b[0];
		this.getData()[1]= b[1];
	}
	
	
	/**
	 * returns the cache data sources
	 * <br>NOT FULLY IMPLEMENTED - only valid for sheet data range data soures
	 * @return
	 */
	public CellRange getCellRange() {
		for (int i= 0; i < subRecs.size(); i++) {
			BiffRec br= (BiffRec) subRecs.get(i);
			if (br.getOpcode()==SXVS) {
				if (((SxVS)br).getSourceType()!=SxVS.TYPE_SHEET) {
					Logger.logErr("SXSTREAMID.getCellRange:  Pivot Table Data Sources other than Sheet are not supported");
					return null;
				}
			} else if (br.getOpcode()==DCONREF) {
				return  ((DConRef)br).getCellRange();
			} else if (br.getOpcode()==DCONNAME) {
				Logger.logErr("SXSTREAMID.getCellRange:  Name sources are not yet supported");
				return null;
			} else if (br.getOpcode()==DCONBIN) {
				Logger.logErr("SXSTREAMID.getCellRange:  Name sources are not yet supported");
				return null;
			}
		}
		return null;
	}
	
	/**
	 * sets the cell range for this pivot cache
	 * @param cr
	 */
	public void setCellRange(CellRange cr) {
		for (int i= 0; i < subRecs.size(); i++) {
			BiffRec br= (BiffRec) subRecs.get(i);
			if (br.getOpcode()==SXVS) {
				if (((SxVS)br).getSourceType()!=SxVS.TYPE_SHEET) {
					Logger.logErr("SXSTREAMID.setCellRange:  Pivot Table Data Sources other than Sheet are not supported");
					return;
				}
			} else if (br.getOpcode()==DCONREF) {
				((DConRef)br).setCellRange(cr);
				return;
			} else if (br.getOpcode()==DCONNAME) {
				Logger.logErr("SXSTREAMID.getCellRange:  Name sources are not yet supported");
				return;
			} else if (br.getOpcode()==DCONBIN) {
				Logger.logErr("SXSTREAMID.getCellRange:  Name sources are not yet supported");
				return;
			}
		}
	}
	
	/**
	 * sets the cell range for this pivot cache
	 * @param cr
	 */
	public void setCellRange(String cr) {
		for (int i= 0; i < subRecs.size(); i++) {
			BiffRec br= (BiffRec) subRecs.get(i);
			if (br.getOpcode()==SXVS) {
				if (((SxVS)br).getSourceType()!=SxVS.TYPE_SHEET) {
					Logger.logErr("SXSTREAMID.setCellRange:  Pivot Table Data Sources other than Sheet are not supported");
					return;
				}
			} else if (br.getOpcode()==DCONREF) {
				((DConRef)br).setCellRange(cr);
				return;
			} else if (br.getOpcode()==DCONNAME) {
				Logger.logErr("SXSTREAMID.getCellRange:  Name sources are not yet supported");
				return;
			} else if (br.getOpcode()==DCONBIN) {
				Logger.logErr("SXSTREAMID.getCellRange:  Name sources are not yet supported");
				return;
			}
		}
	}
	
	/**
	 * creates the basic, default records necessary to define a pivot cache
	 * @param bk
	 * @param ref			string datasource range or named range reference 
	 * @param sheetName		string datasource sheetname where ref is located
	 * @return	arraylist of records
	 */
	public ArrayList addInitialRecords(WorkBook bk, String ref, String sheetName) {
    	ArrayList initialrecs= new ArrayList();
    	int sid= this.getStreamID();
    	SxVS sxvs= (SxVS) SxVS.getPrototype();
    	addInit(initialrecs, sxvs, bk);
    	if (bk.getName(ref)!=null) {
    		// DConName or DConBin
    		Logger.logErr("PivotCache:  Name Data Sources are Not Supported");
    	} else {	// assume it's a regular reference
    		// DConRef
    		DConRef dc= (DConRef) DConRef.getPrototype();
    		int[] rc= ExcelTools.getRangeRowCol(ref);
    		dc.setRange(rc, sheetName);
        	addInit(initialrecs, dc, bk);
    	}
    	// required SxAddl records: stores additional PivotTableView, PivotCache info of a variety of types
    	byte[] b= ByteTools.cLongToLEBytes(sid);      	b= ByteTools.append(new byte[] {0, 0}, b); // add 2 reserved bytes
    	SxAddl sa= SxAddl.getDefaultAddlRecord(SxAddl.ADDL_CLASSES.sxcCache, SxcCache.SxdId.sxd(), b);	//4 bytes sid, 2 bytes reserved
    	addInit(initialrecs, sa, bk);
    	sa= SxAddl.getDefaultAddlRecord(SxAddl.ADDL_CLASSES.sxcCache, SxcCache.SxdVer10Info.sxd(), null);
    	addInit(initialrecs, sa, bk);
    	sa= SxAddl.getDefaultAddlRecord(SxAddl.ADDL_CLASSES.sxcCache, SxcCache.SxdVerSxMacro.sxd(), null);
    	addInit(initialrecs, sa, bk);
    	sa= SxAddl.getDefaultAddlRecord(SxAddl.ADDL_CLASSES.sxcCache, SxcCache.SxdEnd.sxd(), null);
    	addInit(initialrecs, sa, bk);
		return initialrecs;
	}
    /**
     * utility function to properly add a Pivot Table View subrec  
     * @param initialrecs
     * @param rec
     * @param addToInitRecords
     * @param sheet
     */
    private void addInit(ArrayList initialrecs, XLSRecord rec, WorkBook bk) {
    	rec.setWorkBook(bk);
    	initialrecs.add(rec);
    	this.addSubrecord(rec);
    }
}
