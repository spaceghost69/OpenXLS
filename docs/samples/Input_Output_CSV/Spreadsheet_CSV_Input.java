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
/** ------------------------------------------------------------
 * Spreadsheet_CSV_Input.java
 *
 *
 * ------------------------------------------------------------
 */
package docs.samples.Input_Output_CSV;

import com.valkyrlabs.OpenXLS.CellHandle;
import com.valkyrlabs.OpenXLS.RowHandle;
import com.valkyrlabs.OpenXLS.WorkBookHandle;
import com.valkyrlabs.OpenXLS.WorkSheetHandle;
import com.valkyrlabs.toolkit.Logger;

/**
 * This class demonstrates how to load a Spreadsheet from CSV data
 * 
 *
 */
public class Spreadsheet_CSV_Input {

	/** 
	 * Test functionality of reading and writing CSV file to and from
	 * a Spreadsheet.
	 * 
	 * Reads in from a CSV file and Writes to System.out the Spreadsheet as CSV. 
	 * 
	 * @param args
	 */
	public static void main(String[] args) {
		WorkBookHandle book = null;
		try{ // generate a PDF and confirm
			// load the current file and output it to same directory
			book = new WorkBookHandle(System.getProperty("user.dir") + "/docs/samples/Input_CSV/contacts.csv");
			com.valkyrlabs.OpenXLS.util.Logger.log(book.getStats());
			Logger.logInfo("Successfully read CSV into spreadsheet: " +book);
		}catch(Exception e){
			Logger.logErr("Spreadsheet_CSV_Input failed.",e);
		}
		
		// Write out the spreadsheet as CSV -- 
        try {
            WorkSheetHandle wsh = null;
	    	wsh = book.getWorkSheet(0);
	    	StringBuffer arr = new StringBuffer();
            // return WorkBookCommander.returnXMLErrorResponse("PluginSheet.get() failed: "+ex.toString());
            RowHandle[] rwx = wsh.getRows();
            for (int i=0;i<rwx.length;i++) {
                RowHandle r = rwx[i];
                try{
                	CellHandle[] chx = r.getCells();
                	for(int t=0;t<chx.length;t++){
                		arr.append("'"+chx[t].getFormattedStringVal()+"',");
                	}
                	arr.setLength(arr.length()-1);
                	arr.append("\r\n");
                }catch(Exception ex){
                	Logger.logErr("Spreadsheet CSV output failed to fetch row:"+i,ex);
                }
            }
            com.valkyrlabs.OpenXLS.util.Logger.log("WorkBook:" + book + " CSV output: "+  arr.toString());
        }catch(Exception e){
        	Logger.logErr("Spreadsheet CSV output failed: "+e.toString());
        }
	}
}
