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
package docs.samples.AddingCells;

import java.io.BufferedOutputStream;
import java.io.BufferedReader;
import java.io.File;
import java.io.FileOutputStream;
import java.text.DateFormat;
import java.text.SimpleDateFormat;

import com.valkyrlabs.OpenXLS.CellHandle;
import com.valkyrlabs.OpenXLS.DateConverter;
import com.valkyrlabs.OpenXLS.WorkBookHandle;
import com.valkyrlabs.OpenXLS.WorkSheetHandle;
import com.valkyrlabs.formats.XLS.SheetNotFoundException;
import com.valkyrlabs.toolkit.Logger;
import com.valkyrlabs.toolkit.StringTool;

/** This example shows OpenXLS with all the high performance settings
 *  enabled, and optimized for adding 
 * 
 * 
 *  this test uses a bit of memory, be sure to set the Xmx setting on the Java command line
 * 
 * 	ie: -Xms16M -Xmx1032M
 *
 *  ------------------------------------------------------------
 * 
 */
public class HighPerformance{
    
    public static DateFormat in_format = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
    
    public static void main(String[] args){
        
        String wd = System.getProperty("user.dir")+"/docs/samples/AddingCells/";

        for(int z=24000;z<25000;z+=1000){      
        WorkBookHandle bookHandle = new WorkBookHandle();
        
        // bookHandle.setDupeStringMode(WorkBookHandle.SHAREDUPES);
        bookHandle.setStringEncodingMode(WorkBookHandle.STRING_ENCODING_COMPRESSED); // Change to UNICODE if you have eastern strings
        
       // System.setProperty("com.valkyrlabs.OpenXLS.cacheCellHandles","true");
        
        Logger.logInfo("OpenXLS Version: " + WorkBookHandle.getVersion());
        BufferedReader fileReader = null;
        Logger.logInfo("Begin test.");
       
	        try{
	            int recordNum = 0;
		        int sheetNum = 0;
		        int row = 0;
	            
		        WorkSheetHandle sheetHandle;
	            sheetHandle = addWorkSheet(bookHandle, sheetNum);
	            //Logger.logInfo(bookHandle.getStats());
	            
	            // See the valid list of format patterns in the 
	            // API docs for FormatHandle
	            WorkSheetHandle formatSheet = bookHandle.getWorkSheet("Sheet2"); 
	            formatSheet.setHidden(true);
	            
	            // These cells are added to create a format in the workbook...
	            // 
	            CellHandle currencycell = formatSheet.add(new Double(123.234), "A1");
	            currencycell.setFormatPattern("($#,##0);($#,##0)");
	            
	            CellHandle numericcell = formatSheet.add(new Double(123.234), "A2");
	            numericcell.setFormatPattern("0.00");
	            
	            CellHandle datecell = formatSheet.add(new java.util.Date(System.currentTimeMillis()), "A3");
	            
	            int CurrencyFormat = formatSheet.getCell("A1").getFormatId();
	            int NumericFormat  = formatSheet.getCell("A2").getFormatId();
	            int DateFormat  = formatSheet.getCell("A3").getFormatId();
	            
	            Logger.logInfo("Starting adding cells.");
	            String line = "1234	3			4 ZZZZZZZZZZZZ	640	2	6			2005-01-28 00:00:00	8	9	7477747	QA01898388			2005-01-28 00:00:00	2005-01-28 00:00:00		0	0	0	0	0	0	1805000	1805000		1805000		2	0				NL	8	7	SOME ACCOUNT INC	293881	72	AKZO ZZZZZZZZZZZZ				28-Jan-05	783321	802778	99999	1294092184	640	1857520	A\r\n";
	
	            /**
	             *  NOTE:This is a very important setting for performance
	             *  eliminates the lookup and return of a CellHandle for 
	             *  each new Cell added 
	             */
	            sheetHandle.setFastCellAdds(true);
	            
	            for(int t=0;t<z;t++){
	
	               String[] tokens = StringTool.getTokensUsingDelim(line,"\t");
	               if (tokens != null)
	               {
	                   for (int i = 0; i < tokens.length; i++)
	                   {
		                   	if ( tokens[i] != null && recordNum!=0)
		                  	{
		                       switch (i) {
		                       
	                       	   case 10:
	                       	       ;
	                       	   case 17:
	                       	       ;
	                       	   case 18:
	                       	       ;
	                       	   case 34:
	                       	       
		                       	        String dtsr = tokens[i];
		                       	        if(!dtsr.equals("")) {
			                       	        java.sql.Date dtx = new java.sql.Date(in_format.parse(dtsr).getTime());
			                       	        
			                       	        // change the date pattern here...
			                       	        double dr = DateConverter.getXLSDateVal(dtx);
			                       	        sheetHandle.add(new Double(dr),row,i);
		                       	        }	                               
		                       	       break;
	                       	   case 16:
	                       	   case 20:
	                       	   case 21:		                       	       
	                       	   case 22:
	                       	   case 23:
	                       	   case 24:
	                       	   case 25:
	                       	   case 26:
	                       	   case 27:
	                       	   case 29:
	                       	   case 31:
	                       		   if(!tokens[i].equals("")) {
										// allows you to store numbers as numbers in XLS
										Object ob = null;
										try { // try to get it as a number
										    ob = new Double(tokens[i]+t);
										}catch(Exception ex) {
										    try { // try to get it as a number
											    ob = Integer.valueOf(tokens[i]+t);
											}catch(Exception ext) {
											    ob =  tokens[i]+t;
											}
										}
										sheetHandle.add(ob,row,i);
	                       		   }
									break;
	
	                       	   default :
		                       		if(!tokens[i].equals("")) {
										// allows you to store numbers as numbers in XLS
										Object ob = tokens[i]+t;
										try { // try to get it as a number
										    ob = new Double(ob.toString());
										}catch(Exception ex) {
										    try { // try to get it as a number
											    ob = Integer.valueOf(ob.toString());
											}catch(Exception ext) {
											    ;
											}
										}
										sheetHandle.add(ob,row,i);
									}
		                       }
		                  	}
		                	else if (tokens[i] != null)
		                	{
		                   		
		                	    sheetHandle.add(tokens[i],row,i);
		                	}
	                       
	                   }
	               }
	            
	               recordNum++;
	               row++;
	               if (recordNum%1000 == 0)
	               {
	                   Logger.logInfo(recordNum + " Rows Added");
	               }
	               if (recordNum%65000== 0){
	                   row=0;
	                   sheetNum++;
	                   sheetHandle = addWorkSheet(bookHandle, sheetNum);
	                   sheetHandle.setFastCellAdds(true);
	               }
	            }
	        }
	        catch (Exception e){
	            Logger.logErr(e);
	        }finally{
	            File oFile = new File(wd + "fastAddOut_"+z+".xls");
	            
	            try{
	                Logger.logInfo("Begin writing XLS file...");
                    // MUST use a buffered out for writing performance
                    BufferedOutputStream bout = new BufferedOutputStream(new FileOutputStream(oFile));
                    bookHandle.write(bout);
                    bout.flush();
                    bout.close();
	                Logger.logInfo("Done writing XLS file.");
	            }catch (Exception e1){
	                Logger.logErr(e1);
	            }
	            Logger.logInfo("Start reading XLS file.");
	            WorkBookHandle wbh = new WorkBookHandle(wd + "fastAddOut_"+z+".xls");
	            Logger.logInfo("Done reading XLS file.");
	            wbh = null;
	            bookHandle = null;
	            System.gc();
	        }
            Logger.logInfo("End test.");
        }
    }

    /**
     * @param bookHandle
     * @param k
     * @param sheetHandle
     * @return
     * @throws SheetNotFoundException
     */
    private static WorkSheetHandle addWorkSheet(WorkBookHandle bookHandle, int k) throws SheetNotFoundException
    {
        WorkSheetHandle sheetHandle = null;
        try
        {
            sheetHandle = bookHandle.getWorkSheet(k);
        }
        catch (com.valkyrlabs.formats.XLS.SheetNotFoundException
        {
            sheetHandle = bookHandle.createWorkSheet("sheet" + k);
             
        }
        return sheetHandle;
    }
}
