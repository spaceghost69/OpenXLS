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
package docs.samples.Formats;

import com.valkyrlabs.OpenXLS.*;

import java.awt.Color;
import java.io.BufferedOutputStream;
import java.io.FileOutputStream;

import com.valkyrlabs.OpenXLS.CellHandle;
import com.valkyrlabs.OpenXLS.CellRange;
import com.valkyrlabs.OpenXLS.FormatHandle;
import com.valkyrlabs.OpenXLS.WorkBookHandle;
import com.valkyrlabs.formats.XLS.FormatConstants;
import com.valkyrlabs.toolkit.Logger;

/**
 *  Demonstrates creating borders on Cell Ranges
 *   
 *
 */
public class testBorders {
	

	String wd = System.getProperty("user.dir")+"/docs/samples/Formats/";
	
	/**
	 * Demonstrates creating borders on Cell Ranges
	 * Jan 19, 2010
	 * @param args
	 */
	public static void main(String[] args){
		testBorders tb = new testBorders();
		tb.testit();
	}
	
	/**
	 * tests various border types and cell sides
	 * ------------------------------------------------------------
	 * 
	 */
	public void testit() {
		Logger.logInfo("====================== TESTING borders on cells ======================");
        WorkBookHandle tbo = new WorkBookHandle();
        try{
			WorkSheetHandle sheet = tbo.getWorkSheet(0);
			
            int[] coords = {12,1,12,5};
            CellRange range = new CellRange(sheet, coords, true);
            
            CellRange range2= new CellRange("Sheet1!B2:Sheet1!C10", tbo, true);
            range2.setBorder(2,FormatConstants.BORDER_THIN, Color.blue);
            
            // set top and bottom
            FormatHandle myfmthandle = new FormatHandle(tbo);
            myfmthandle.addCellRange(range);
	        myfmthandle.setTopBorderLineStyle(FormatHandle.BORDER_DOUBLE);
	        myfmthandle.setBottomBorderLineStyle(FormatHandle.BORDER_THICK);
	        
            // set sides
	        int[] coords2 = {5,4,5,8};
	        CellRange range3 = new CellRange(sheet, coords2, true);
            FormatHandle myfmthandle2 = new FormatHandle(tbo);
            myfmthandle2.addCellRange(range3);
	        myfmthandle2.setBorderLeftColor(Color.red);
            myfmthandle2.setLeftBorderLineStyle(FormatHandle.BORDER_DASH_DOT_DOT);
            
            myfmthandle2.setBorderRightColor(Color.blue);
            myfmthandle2.setRightBorderLineStyle(FormatHandle.BORDER_DOUBLE);
	        
            // ok, test not clobbering
            CellRange range4 = new CellRange(sheet, coords2, true);
            
            CellHandle cell0 = range4.getCells()[0];
           
            FormatHandle clobberfmt = cell0.getFormatHandle();
            clobberfmt.setCellBackgroundColor(Color.lightGray);
            clobberfmt.setUnderlined(true);
            
            cell0.setVal("hello world!");
            
            
		}catch(Exception ex) {
			Logger.logErr("testCellBorder failed:" + ex.toString());
		}
		testWrite(tbo);
		
		Logger.logInfo("====================== DONE TESTING borders on cells ======================");
        	
	}

    public void testWrite(WorkBookHandle book){
        try{
      	    java.io.File f = new java.io.File(wd + "testBorders_out.xls");
            FileOutputStream fos = new FileOutputStream(f);
            BufferedOutputStream bbout = new BufferedOutputStream(fos);
            book.write(bbout);
            bbout.flush();
		    fos.close();
      	} catch (java.io.IOException e){Logger.logInfo("IOException in Tester.  "+e);}  
    }	
}