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

import java.io.UnsupportedEncodingException;
import java.util.Arrays;
import java.util.Locale;

import com.valkyrlabs.OpenXLS.DateConverter;
import com.valkyrlabs.formats.XLS.formulas.DateTimeCalculator;
import com.valkyrlabs.toolkit.ByteTools;
import com.valkyrlabs.toolkit.Logger;

/**
 * The SXAddl record specifies additional information for a PivotTable view,
 * PivotCache, or query table. The current class and full type of this record
 * are specified by the hdr field which determines the contents of the data
 * field.
 * 
 * hdr (6 bytes): An SXAddlHdr structure that specifies header information for
 * an SXAddl record. data (variable): A variable-size field that contains data
 * specific to the full record type of the SXAddl record.
 * 
 * 
 * The SXAddlHdr structure specifies header information for an SXAddl record.
 * frtHeaderOld (4 bytes): An FrtHeaderOld. The frtHeaderOld.rt field MUST be
 * 0x0864. sxc (1 byte): An unsigned integer that specifies the current class.
 * See class for details. sxd (1 byte): An unsigned integer that specifies the
 * type of record contained in the data field of the containing SXAddl record.
 * See class for details
 * 
 * 
 *         "http://www.extentech.com">Extentech Inc.</a>
 * 
 */
public class SxAddl extends XLSRecord implements XLSConstants {
	/**
	 * serialVersionUID
	 */
	private static final long serialVersionUID = 2639291289806138985L;
	private short sxc, sxd;

	enum ADDL_CLASSES {
		sxcView(0), sxcField(1), sxcHierarchy(2), sxcCache(3), /* 3 */
		sxcCacheField(4), sxcQsi(5), sxcQuery(6), sxcGrpLevel(7), sxcGroup(8), sxcCacheItem(
				9), /* 9 */
		sxcSxrule(0xC), sxcSxfilt(0xD), sxcSxdh(0x10), sxcAutoSort(0x12), sxcSxmgs(
				0x13), sxcSxmg(0x14), sxcField12(0x17), sxcSxcondfmts(0x1A), sxcSxcondfmt(
				0x1B), sxcSxfilters12(0x1C), sxcSxfilter12(0x1D);
		private final short cls;

		ADDL_CLASSES(int cls) {
			this.cls = (short) cls;
		}

		public short sxd() {
			return cls;
		}

		public static ADDL_CLASSES get(int cls) {
			for (ADDL_CLASSES c : values()) {
				if (c.cls == cls)
					return c;
			}
			return null;
		}
	}

	/** sxc= 3 **/
	enum SxcCache {
		SxdId(0), // SXAddl_SXCCache_SXDId
		SxdVerUpdInv(1), // SXAddl_SXCCache_SXDVerUpdInv
		SxdVer10Info(2), // SXAddl_SXCCache_SXDVer10Info
		SxdVerSxMacro(0x18), // SXAddl_SXCCache_SXDVerSXMacro 0x18 (24)
		SxdInvRefreshReal(0x34), // SXAddl_SXCCache_SXDInvRefreshReal 0x34
		SxdInfo12(0x41), // SXAddl_SXCCache_SXDInfo12 0x41
		SxdEnd(-1); // SXAddl_SXCCache_SXDEnd 0xFF
		private final short sxd;

		SxcCache(int sxd) {
			this.sxd = (short) sxd;
		}

		public short sxd() {
			return sxd;
		}

		public static SxcCache lookup(int record) {
			for (SxcCache c : values()) {
				if (c.sxd == record)
					return c;
			}
			return null;
		}
	}

	/** sxc= 0 **/
	enum SxcView {
		sxdId(0), // SXAddl_SXCView_SXDId
		sxdVerUpdInv(1), // SXAddl_SXCView_SXDVerUpdInv
		sxdVer10Info(2), // SXAddl_SXCView_SXDVer10Info
		sxdCalcMember(3), // SXAddl_SXCView_SXDCalcMember
		sxdCalcMemString(0xA), // SXAddl_SXCView_SXDCalcMemString 0xA
		sxdVer12Info(0x19), // SXAddl_SXCView_SXDVer12Info 0x19
		sxdTableStyleClient(0x1E), // SXAddl_SXCView_SXDTableStyleClient 0x1E
									// (30)
		sxdCompactRwHdr(0x21), // SXAddl_SXCView_SXDCompactRwHdr 0x21
		sxdCompactColHdr(0x22), // SXAddl_SXCView_SXDCompactColHdr 0x22
		sxdSxpiIvmb(0x26), // SXAddl_SXCView_SXDSXPIIvmb 0x26
		sxdEnd(-1); // SXAddl_SXCView_SXDEnd 0xFF
		private final short sxd;

		SxcView(int sxd) {
			this.sxd = (short) sxd;
		}

		public short sxd() {
			return sxd;
		}

		public static SxcView lookup(int record) {
			for (SxcView c : values()) {
				if (c.sxd == record)
					return c;
			}
			return null;
		}
	}

	/** sxd= 0x17 **/
	enum SxcField12 {
		sxdId(0), sxdVerUpdInv(1), sxdMemberCaption(0x11), sxdVer12Info(0x19), sxdIsxth(
				0x1C), sxdAutoshow(0x37), sxdEnd(-1);
		private final short sxd;

		SxcField12(int sxd) {
			this.sxd = (short) sxd;
		}

		public short sxd() {
			return sxd;
		}

		public static SxcField12 lookup(int record) {
			for (SxcField12 c : values()) {
				if (c.sxd == record)
					return c;
			}
			return null;
		}
	}

	public void init() {
		super.init();
		sxc = this.getData()[4]; // class: see addlclass
		sxd = this.getData()[5];
		int len = this.getData().length;
		/*
		 * notes: If the value of the hdr.sxc field of SXAddl is 0x09 and the
		 * value of the hdr.sxd field of SXAddl is 0xFF, then the current class
		 * is specified by SxcCacheField class and the full record type is
		 * SXAddl_SXCCacheItem_SXDEnd. Classes can be nested inside other
		 * classes in a hierarchical manner
		 */
		switch (ADDL_CLASSES.get(sxc)) {
		case sxcView: /* 0 */
			SxcView record = SxcView.lookup(sxd);
			switch (record) {
			case sxdId:
				// An SXAddl_SXString structure that specifies the PivotTable
				// view that this SxcView class applies to.
				// The corresponding SxView record of this PivotTable view is
				// the SxView record, in this Worksheet substream,
				// with its stTable field equal to the value of this field.
				// SXADDL_sxcView: record=sxdId data:[11, 0, 0, 0, 0, 0, 11, 0,
				// 0, 80, 105, 118, 111, 116, 84, 97, 98, 108, 101, 50]
				// cchTotal 4 bytes -- if multiple segments (for strings > 255)
				// will be 0
				// reserved 2 bytes
				// String-- cch-2 bytes, encoding-1 byte
				if (len > 6) {
					short cch = ByteTools.readShort(this.getData()[6],
							this.getData()[7]);
					if (cch > 0) { // otherwise it's a multiple segment
						cch = ByteTools.readShort(this.getData()[12],
								this.getData()[13]);
						short encoding = this.getData()[14];
						byte[] tmp = this
								.getBytesAt(15, (cch) * (encoding + 1));
						String name = null;
						try {
							if (encoding == 0)
								name = new String(tmp, DEFAULTENCODING);
							else
								name = new String(tmp, UNICODEENCODING);
						} catch (UnsupportedEncodingException e) {
							Logger.logInfo("encoding PivotTable caption name in Sxvd: "
									+ e);
						}
						if (DEBUGLEVEL > 3)
							Logger.logInfo("SXADDL_sxcView: record=" + record
									+ " name: " + name);
					} else if (DEBUGLEVEL > 3)
						Logger.logInfo("SXADDL_sxcView: record=" + record
								+ " name: MULTIPLESEGMENTS");
				} else if (DEBUGLEVEL > 3)
					Logger.logInfo("SXADDL_sxcView: record=" + record
							+ " name: null");
				break;
			case sxdVer10Info:
			case sxdTableStyleClient:
			case sxdVerUpdInv:
				if (DEBUGLEVEL > 3)
					Logger.logInfo("SXADDL_sxcView: record=" + record
							+ " data:"
							+ Arrays.toString(this.getBytesAt(6, len - 6)));
				break;
			}
			break;
		case sxcCache: /* 3 */
			SxcCache crec = SxcCache.lookup(sxd);
			switch (crec) {
				case SxdVer10Info:
					byte verLastRefresh= this.getByteAt(16);
					byte verRefreshMin= this.getByteAt(17);
					double lastdate= ByteTools.eightBytetoLEDouble(this.getBytesAt(18, 8));
					java.util.Date ld= DateConverter.getDateFromNumber(lastdate);
					if (DEBUGLEVEL > 3){
						java.text.DateFormat dateFormatter= java.text.DateFormat.getDateInstance(java.text.DateFormat.DEFAULT, Locale.getDefault());
						Logger.logInfo("SXADDL_sxcCache: record=" + crec + 
								" lastDate:" + dateFormatter.format(ld) + " verLast:" + verLastRefresh 
								+ " verMin:" + verRefreshMin);
					}
					break;
				default:
					if (DEBUGLEVEL > 3)
						Logger.logInfo("SXADDL_sxcCache: record=" + crec + " data:"
								+ Arrays.toString(this.getBytesAt(6, len - 6)));
					break;
			}
			
			break;
		case sxcField12:
			SxcField12 srec = SxcField12.lookup(sxd);
			if (DEBUGLEVEL > 3)
				Logger.logInfo("SXADDL_sxcField12: record=" + srec + " data:"
						+ Arrays.toString(this.getBytesAt(6, len - 6)));
			break;
		case sxcField:
		case sxcHierarchy:
		case sxcCacheField:
		case sxcQsi:
		case sxcQuery:
		case sxcGrpLevel:
		case sxcGroup:
		case sxcCacheItem:
		case sxcSxrule:
		case sxcSxfilt:
		case sxcSxdh:
		case sxcAutoSort:
		case sxcSxmgs:
		case sxcSxmg:
		case sxcSxcondfmts:
		case sxcSxcondfmt:
		case sxcSxfilters12:
		case sxcSxfilter12:
			if (DEBUGLEVEL > 3)
				Logger.logInfo("SXADDL: hdr: " + " sxc:" + sxc + " sxd:" + sxd
						+ " data:"
						+ Arrays.toString(this.getBytesAt(6, len - 6)));
			break;
		}
	}

	/**
	 * creates a SxAddl record for the desired class and record id
	 * 
	 * @param cls
	 *            int class one of ADDL_CLASSES enum
	 * @param recordid
	 *            desired record in class
	 * @param dara
	 *            if not null,specifies the data for the class. if null, the
	 *            default data will be used
	 * @return SxAddl record
	 */
	public static SxAddl getDefaultAddlRecord(SxAddl.ADDL_CLASSES cls,
			int recordid, byte[] data) {
		SxAddl sxa = new SxAddl();
		sxa.setOpcode(SXADDL);
		byte[] newData = new byte[6];
		newData[0] = 100;
		newData[1] = 8;
		newData[4] = (byte) cls.ordinal();
		newData[5] = (byte) recordid;

		if (data == null) { // if !null, use passed in data for record creation
							// and return; otherwise create default data for
							// record
			switch (cls) {
			case sxcView: /* 0 */
				SxcView record = SxcView.lookup(recordid);
				switch (record) {
				case sxdId:
				case sxdTableStyleClient:
				case sxdVerUpdInv:
					break;
				case sxdVer10Info:
					data= new byte[] { 1, 0x41, 0, 0, 0, 0}; // common flags
					break;
				case sxdEnd:
					data = new byte[] { 0, 0, 0, 0, 0, 0 };
					break;
				}
				break;
			case sxcCache: /* 3 */
				SxcCache crec = SxcCache.lookup(recordid);
				switch (crec) {
				case SxdId: // pivot cache stream id
					data = new byte[] { 1, 0, 0, 0, 0, 0 };
					break;
				case SxdVer10Info:					
					data = new byte[] { 0, 0, 0, 0, 0, 0, -1, -1, -1, -1, /* reserved, citmGhostMax */
							3,	/* ver last saved -- 0 or 3 */ 
							0};	/* ver min */ 
					//date last refreshed: 8 bytes 	 
					//reserved: 0, 0 };/* reserved 2 bytes*/
					double d= DateConverter.getXLSDateVal(new java.util.Date());
					byte[] dates= ByteTools.doubleToLEByteArray(d);
					data= ByteTools.append(dates, data);
					data= ByteTools.append(new byte[] {0,0}, data);
					break;
				case SxdVerSxMacro: // DataFunctionalityLevel
					data = new byte[] { 1, 0, 0, 0, 0, 0 };
					break;
				case SxdEnd:
					data = new byte[] { 0, 0, 0, 0, 0, 0 };
					break;
				}
			}
		}
		newData = ByteTools.append(data, newData);
		sxa.setData(newData);
		sxa.init();
		return sxa;
	}

	/**
	 * for SXADDL_SxView_SxDID record this sets the view name (matches table
	 * name in Sxview)
	 * 
	 * @param viewName
	 */
	public void setViewName(String viewName) {
		if (sxc != 0 && sxd != 0)
			Logger.logErr("Incorrect SXADDL_ record for view name");

		byte[] data = new byte[14];
		System.arraycopy(this.getData(), 0, data, 0, 5);
		byte[] strbytes = null;
		try {
			strbytes = viewName.getBytes(DEFAULTENCODING);
		} catch (UnsupportedEncodingException e) {
			Logger.logInfo("encoding pivot view name in SXADDL: " + e);
		}

		// update the lengths:
		short cch = (short) strbytes.length;
		byte[] nm = ByteTools.shortToLEBytes(cch);
		data[6] = nm[0];
		data[7] = nm[1];
		data[12] = nm[0];
		data[13] = nm[1];

		// now append variable-length string data
		byte[] newrgch = new byte[cch + 1]; // account for encoding bytes
		System.arraycopy(strbytes, 0, newrgch, 1, cch);

		data = ByteTools.append(newrgch, data);
		this.setData(data);
	}

	/**
	 * returns the class of this SXADDL_ record
	 * 
	 * @return ADDL_CLASSES instance
	 */
	public ADDL_CLASSES getADDlClass() {
		return ADDL_CLASSES.get(sxc);
	}

	/**
	 * returns the class which matches the class/record id of this SXADDL_
	 * record
	 * 
	 * @return
	 */
	public Object getRecordId() {
		switch (ADDL_CLASSES.get(sxc)) {
		case sxcView: /* 0 */
			return SxcView.lookup(sxd);
		case sxcCache:
			return SxcCache.lookup(sxd);
		case sxcField12:
			return SxcField12.lookup(sxd);
		}
		return null;
	}
}
