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
package docs.samples.NamedRanges;


import com.valkyrlabs.OpenXLS.*;
import com.valkyrlabs.formats.XLS.*;
import com.valkyrlabs.toolkit.Logger;

import java.io.*;


/**

    This Class Demonstrates the basic functionality of of OpenXLS.

 */
public class TestNamedRanges{

    public static void main(String[] args){
        testnames t = new testnames();
		String s = "Test Successful.";
		
		t.testit(s);
    }
}

/** Test the handling of and modification of Named Ranges
*/
class testnames{
	String wd = System.getProperty("user.dir")+"/docs/samples/NamedRanges/";

    public void testit(String argstr){
        WorkBookHandle tbo = new WorkBookHandle(wd + "testNames.xls");
        try{
            WorkSheetHandle sheet1 = tbo.getWorkSheet("Sheet1");

            String ht="E3:E10";
            for(int t = 3; t<=10;t++){
                try{
                    sheet1.add(new Float(t*67.5),"E" + t);
                }catch(Exception e){;}
            }
           
           // this will throw a CellNotFoundException if the name is not found.
            NameHandle nand = tbo.getNamedRange("nametest4");
            CellHandle[] ch = nand.getCells();
            for(int x = 0;x<ch.length;x++){
                ch[x].setVal(123 * x);
                ch[x].setFontColor(FormatHandle.PaleBlue);
            }
            com.valkyrlabs.OpenXLS.util.Logger.log(nand.getName());
            
            NameHandle nand2 = tbo.getNamedRange("nametest7");
            CellHandle[] ch1 = nand2.getCells();
            for(int x = 0;x<ch1.length;x++){
                com.valkyrlabs.OpenXLS.util.Logger.log(ch1[x].getWorkSheetName() +":"+ ch1[x].getCellAddress() +":"+ ch1[x].getVal());
            }
            nand2.setName("IMPORTANT");
            com.valkyrlabs.OpenXLS.util.Logger.log(nand2.getName());            
            nand2.setLocation("A1:B15");     
            
            NameHandle nand3 = tbo.getNamedRange("nametest10");
            CellHandle[] ch2 = nand3.getCells();
            
            CellHandle[] cx = nand3.getCells();
            
//            sheet1.add("Ken", cx[4].getRowNum(), cx[4].getColNum());
            
            
            for(int x = 0;x<ch2.length;x++){
                com.valkyrlabs.OpenXLS.util.Logger.log(ch2[x].getWorkSheetName() +":"+ ch2[x].getCellAddress() +":"+ ch2[x].getVal());
            }
            com.valkyrlabs.OpenXLS.util.Logger.log(nand3.getName());        
            nand3.setName("URGENTCELLS");
            nand3.setLocation("D15");    

            // Create 2 named ranges for Constant values (formulas)
            
            // create a 'true' constant
            NameHandle nh = tbo.createNamedRange("testtrue", "=true");
            
            // create a 'false' constant
            NameHandle fn = tbo.createNamedRange("testfalse", "=false");

            
			// Create new NameHandle from a CellRange
			CellRange range = new CellRange( "Sheet1!D8:D13", tbo);
			NameHandle newname = new NameHandle("NEWNAME",range);	
			try{
				CellRange[] ranges = newname.getCellRanges();
				for(int t=0;t<ranges.length;t++){
					com.valkyrlabs.OpenXLS.util.Logger.log("Got new range:" + ranges[t].toString());
				}
			}catch(Exception e){
				System.err.println("Problem creating new Named Range:" + newname.toString());
			}
        }catch(CellNotFoundException e){com.valkyrlabs.OpenXLS.util.Logger.log(e);}
        catch(SheetNotFoundException e){com.valkyrlabs.OpenXLS.util.Logger.log(e);}
        testWrite(tbo);
    }

    public void testWrite(WorkBookHandle b){
        try{
      	    java.io.File f = new java.io.File(wd + "testNamesOutput.xls");
            FileOutputStream fos = new FileOutputStream(f);
            BufferedOutputStream bbout = new BufferedOutputStream(fos);
            b.write(bbout);
            bbout.flush();
		    fos.close();
      	} catch (java.io.IOException e){Logger.logInfo("IOException in Tester.  "+e);}  
    }

}