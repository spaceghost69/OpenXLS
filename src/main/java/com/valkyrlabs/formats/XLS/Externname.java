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

import com.valkyrlabs.toolkit.ByteTools;
import com.valkyrlabs.toolkit.Logger;

// TODO: Handle other types of External Names besides Add-Ins

/** <b>Externname: External Name Record (17h)</b><br>

	<p><pre>
	for add-in formulas (1 for each add-in formula)
	
	   offset  name        size    contents
	---    
	0       Op Flags 		2       Always 00 for Add-ins
	2		Not used		4
	6		Name			var.	Unicode formula name (1st 2 bytes are length)
	var.	#REF Err Code	4		always 02 00 1C 17  (2 0 28 23)
	
	for external names
	   offset  name        size    contents
	---    
	0       Op Flags 		2       Always 00 for Add-ins
	2		INDEX			2		One-based index to EXTERNSHEET 
	4		Not used		2
	6		Name			var.	External name	
	var.	Formula data	var.	RPN Token Array
	</pre></p>
   
   @see WorkBook
   @see Boundsheet
   @see Externsheet
   
*/
public final class Externname extends XLSRecord {    
    /** 
	* serialVersionUID
	*/
	private static final long serialVersionUID = -7153354861666069899L;

	public void setWorkBook(WorkBook bk) {
      super.setWorkBook(bk);
    }
      
   	public void init(){
       super.init();
       short externnametype= ByteTools.readShort(this.getByteAt(0), this.getByteAt(1));
       if (externnametype==0) {// add-in
           // read in length of external name
           short len= ByteTools.readShort(this.getByteAt(6),this.getByteAt(7));
           byte[] b= this.getBytesAt(8, len);
           String s= new String(b);
           getWorkBook().addExternalName(s);	// store external names in workbook
       }
	}
   	static final byte[] ENDOFRECORDBYTES= {0x2, 0x0, 0x1C, 0x17};
   	static final int STRINGLENPOS= 6;
   	static final int STRINGPOS= 8;
   	static final int STATICPORTIONSIZE= 12; // header=6, endofrecord=4, strlength=2
   	
   	// TODO: Finish for other types of external names (???)
	protected static XLSRecord getPrototype(String s){
		Externname x = new Externname();
		int len= s.length();		
		x.setLength((short)(STATICPORTIONSIZE + len));	
		x.setOpcode(EXTERNNAME);
		byte[] dta = new byte[STATICPORTIONSIZE + len];
		// write string		
		try {
			// write length 
			byte[] slen= ByteTools.shortToLEBytes((short) len);
			dta[STRINGLENPOS]= slen[0];
			dta[STRINGLENPOS+1]= slen[1];
			// write string
			byte[] bts = s.getBytes();
			System.arraycopy(bts, 0, dta, STRINGPOS, bts.length);
			// write end of record 
			System.arraycopy(ENDOFRECORDBYTES, 0, dta, STRINGPOS + len, ENDOFRECORDBYTES.length);
			
			x.setData(dta);
			x.originalsize = STATICPORTIONSIZE + len;
		}
		catch (Exception e) { Logger.logWarn("Exception excountered writing Externname: " + e.toString()); }
		return x;
//            cch = bts.length/2;
 //           byte[] newbytes = new byte[cch+3];
//            byte[] cchx = com.valkyrlabs.toolkit.ByteTools.shortToLEBytes((short)cch);
 //           newbytes[0] = cchx[0];
  //          newbytes[1] = cchx[1];
   //         newbytes[2] = 0x1;
    //        System.arraycopy(bts,0,newbytes,3, bts.length);
     //       this.setData(newbytes);
	}
   	
	public String[] getExternalNames() {
		return getWorkBook().getExternalNames();
	}

	public String getExternalName(int t) {
		return getWorkBook().getExternalName(t);
	}
}
