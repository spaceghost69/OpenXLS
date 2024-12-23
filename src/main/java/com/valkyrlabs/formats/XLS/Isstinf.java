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

import java.io.Serializable;

import com.valkyrlabs.toolkit.ByteTools;


/** ISSTINF Structure:
    
    offset  name        size    contents
    ---    
    0       ib          4       Stream position into SST record
    4       cb          2       Offset into SST to where 'bucket' begins
    6       (Reserved)  2       Must be zero.
    </p></pre>
    
    * @see SST
    * @see EXTSST
*/
public class Isstinf implements Serializable
{
    /** 
	* serialVersionUID
	*/
	private static final long serialVersionUID = 989277769893893586L;
	int ib;
    short cb;
    short myst;
    
    public Isstinf(byte[] dta){
        this.ib = ByteTools.readInt(dta[0],dta[1],dta[2],dta[3]);
        this.cb = ByteTools.readShort(dta[4],dta[5]);
        this.myst = ByteTools.readShort(dta[6],dta[7]);
    }
}