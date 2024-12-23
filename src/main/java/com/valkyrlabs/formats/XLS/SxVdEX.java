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

/**
 * SxVDEx	0x100
 * 
 * The SXVDEx record specifies extended pivot field properties.
 * 
    A - fShowAllItems (1 bit): A bit that specifies whether to show all pivot items for this pivot field, 
    	including pivot items that do not currently exist in the source data. The value MUST be 0 for an OLAP PivotTable view. 
    	MUST be a value from the following table:
    	Value		  	Meaning
    	0x0			    Specifies that all pivot items are not displayed.
    	0x1			    Specifies that all pivot items are displayed.

    B - fDragToRow (1 bit): A bit that specifies whether this pivot field can be placed on the row axis. This value MUST be ignored for an OLAP PivotTable view. 
    	MUST be a value from the following table:
    	Value	    Meaning
    	0x0		    Specifies that the user is prevented from placing this pivot field on the row axis.
    	0x1		    Specifies that the user is not prevented from placing this pivot field on the row axis.

    C - fDragToColumn (1 bit): A bit that specifies whether this pivot field can be placed on the column axis. This value MUST be ignored for an OLAP PivotTable view. MUST be a value from the following table:
    	Value	    Meaning
    	0x0		    Specifies that the user is prevented from placing this pivot field on the column axis.
    	0x1		    Specifies that the user is not prevented from placing this pivot field on the column axis.

    D - fDragToPage (1 bit): A bit that specifies whether this pivot field can be placed on the page axis. This value MUST be ignored for an OLAP PivotTable view. MUST be a value from the following table:
    	Value	    Meaning
    	0x0		    Specifies that the user is prevented from placing this pivot field on the page axis.
    	0x1		    Specifies that the user is not prevented from placing this pivot field on the page axis.

    E - fDragToHide (1 bit): A bit that specifies whether this pivot field can be removed from the PivotTable view. This value MUST be ignored for an OLAP PivotTable view. MUST be a value from the following table:
    	Value	    Meaning
    	0x0		    Specifies that the user is prevented from removing this pivot field from the PivotTable view.
    	0x1		    Specifies that the user is not prevented from removing this pivot field from the PivotTable view.

    F - fNotDragToData (1 bit): A bit that specifies whether this pivot field can be placed on the data axis. This value MUST be ignored for an OLAP PivotTable view. MUST be a value from the following table:
    	Value	    Meaning
    	0x0		    Specifies that the user is not prevented from placing this pivot field on the data axis.
    	0x1		    Specifies that the user is prevented from placing this pivot field on the data axis.

    G - reserved1 (1 bit): MUST be zero, and MUST be ignored.

    H - fServerBased (1 bit): A bit that specifies whether this pivot field is server-based when on the page axis. For more information, see Source Data. 
    	A value of 1 specifies that this pivot field is a server-based pivot field.

    MUST be 1 if and only if the value of the fServerBased field of the SXFDB record of the associated cache field of this pivot field is 1.

    I - reserved2 (1 bit): MUST be zero, and MUST be ignored.

    J - fAutoSort (1 bit): A bit that specifies whether AutoSort will be applied to this pivot field. For more information, see Pivot Field Sorting.

    K - fAscendSort (1 bit): A bit that specifies whether any AutoSort applied to this pivot field will sort in ascending order. MUST be a value from the following table:
    	Value		    Meaning
    	0x0			    Sort in descending order.
    	0x1			    Sort in ascending order.

    L - fAutoShow (1 bit): A bit that specifies whether an AutoShowfilter is applied to this pivot field. For more information, see Simple Filters.

    M - fTopAutoShow (1 bit): A bit that specifies whether any AutoShow filter applied to this pivot field shows the top-ranked or bottom-ranked values. For more information, see Simple Filters. MUST be a value from the following table:
    	Value		    Meaning
    	0x0			    Any AutoShow filter applied to this pivot field shows the bottom-ranked values.
    	0x1			    Any AutoShow filter applied to this pivot field shows the top-ranked values.

    N - fCalculatedField (1 bit): A bit that specifies whether this pivot field is a calculated field. A value of 1 specifies that this pivot field is a calculated field.

    MUST be 1 if and only if the value of the fCalculatedField field of the SXFDB record of the cache field associated with this pivot field is 1.

    O - fPageBreaksBetweenItems (1 bit): A bit that specifies whether a page break (2) is inserted after each pivot item when the PivotTable is printed.

    P - fHideNewItems (1 bit): A bit that specifies whether new pivot items that appear after a refresh are hidden by default. This value MUST be equal to 0 for a non-OLAP PivotTable view.
    	Value	    Meaning
    	0x0		    New pivot items are shown by default.
    	0x1		    New pivot items are hidden by default.
    	
    reserved3 (5 bits): MUST be zero, and MUST be ignored.

    Q - fOutline (1 bit): A bit that specifies whether this pivot field is in outline form. For more information, see PivotTable layout.

    R - fInsertBlankRow (1 bit): A bit that specifies whether to insert a blank row after each pivot item.

    S - fSubtotalAtTop (1 bit): A bit that specifies whether subtotals are displayed at the top of the group when the fOutline field is equal to 1. For more information, see PivotTable layout.

    citmAutoShow (8 bits): An unsigned integer that specifies the number of pivot items to show when the fAutoShow field is equal to 1. 
    The value MUST be greater than or equal to 1 and less than or equal to 255.

    isxdiAutoSort (2 bytes): A signed integer that specifies the data item that AutoSort uses when the fAutoSort field is equal to 1. If the value of the fAutoSort field is one, 
    the value MUST be greater than or equal to zero and less than the count of SXDI records. MUST be a value from the following table:
    Value    	    Meaning
    -1			    Specifies that the values of the pivot items themselves are used.
    Greater than or equal to zero	Specifies a data item index, as specified in Data Items, of the data item that is used.

    isxdiAutoShow (2 bytes): A signed integer that specifies the data item that AutoShow ranks by when the fAutoShow field is equal to 1. 
    For more information, see Simple Filters. If the value of the fAutoShow field is 1, this value MUST be greater than or equal to zero and less than the count of SXDI records. 
    MUST be a value from the following table:
    Value		    Meaning
    -1			    AutoShow is not enabled for this pivot field.
    Greater than or equal to zero	    Specifies a data item index, as specified in Data Items, of the data item that is used.

    ifmt (2 bytes): An IFmt structure that specifies the number format of this pivot field.

    subName (variable): An optional SXVDEx_Opt structure that specifies the name of the aggregate function used to calculate this pivot field's subtotals. SHOULD<124> be present.

 * 
 * 
 *
 */
public class SxVdEX extends XLSRecord implements XLSConstants {
    /** 
	* serialVersionUID
	*/
	private static final long serialVersionUID = 2639291289806138985L;
	private short citmAutoShow, isxdiAutoSort, isxdiAutoShow, ifmt; 
	public void init(){
        super.init();
        // TODO: flags
        citmAutoShow = this.getByteAt(4);
        isxdiAutoSort= ByteTools.readShort(this.getByteAt(4), this.getByteAt(5));
        isxdiAutoShow= ByteTools.readShort(this.getByteAt(6), this.getByteAt(7));
        ifmt= ByteTools.readShort(this.getByteAt(8), this.getByteAt(9));
        // TODO: subName (variable): An optional SXVDEx_Opt structure that specifies the name of the aggregate function used to calculate this pivot field's subtotals. SHOULD<124> be present.
        
		if (DEBUGLEVEL > 3) Logger.logInfo("SXVDEX - citmAutoShow:" + citmAutoShow + " isxdiAutoSort:" + isxdiAutoSort + " isxdoAutoShow:" + isxdiAutoShow + " ifmt:" + ifmt);
	}	
    private byte[] PROTOTYPE_BYTES = new byte[]   { // default configuration
    		30, 20, 0, 10, -1, -1, -1, -1, 0, 0, -1, -1, 0, 0, 0, 0, 0, 0, 0, 0
    };
    public static XLSRecord getPrototype() {
    	SxVdEX sv= new SxVdEX();
    	sv.setOpcode(SXVDEX);
    	sv.setData(sv.PROTOTYPE_BYTES);
    	sv.init();
    	return sv;
    }
}
