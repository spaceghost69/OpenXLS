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
/** 
 * TempFileManager.java
 *
 *
 * Feb 27, 2012
 *
 * 
 */
package com.valkyrlabs.toolkit;

import java.io.File;
import java.io.IOException;
import java.util.Properties;

/** 
 *
 */
public class DefaultTempFileGeneratorImpl implements TempFileGenerator{

	
	public File createTempFile(String prefix, String extension)
	throws IOException{
		String tmpdir = System.getProperty("java.io.tmpdir");
		String lastchar = tmpdir.substring(tmpdir.length()-1, tmpdir.length());
		if (!lastchar.equals(File.separator)){
			tmpdir = tmpdir + File.separator;
		}
		tmpdir += "extentech/";
		File target = null;
		try{
			File tdir = new File(tmpdir);
			if(!tdir.exists()){
				tdir.mkdirs();
			}
			tdir.deleteOnExit();
			
		    target = File.createTempFile(prefix, extension, tdir);
		}catch(Exception e){
			Logger.logInfo("Could not access temp dir: " + tmpdir);// could not create the temp folder fallback to unspecified temp file
			target = File.createTempFile(prefix, extension);
		}
	    
	    target.deleteOnExit();
	   //	target.delete(); // triggers the deleteonexit
	    return target;
	}

}
