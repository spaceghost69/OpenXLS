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
package com.valkyrlabs.formats.XLS.formulas;

import com.valkyrlabs.formats.XLS.Formula;
import com.valkyrlabs.toolkit.Logger;

import java.lang.reflect.Array;


/*
   Ptg that is a Less than operand
   
   Evaluates to TRUE if the second operand is less than the top 
   operand, otherwise FALSE
    
 * @see Ptg
 * @see Formula

    
*/
public class PtgLT extends GenericPtg implements Ptg {
    /**
     * serialVersionUID
     */
    private static final long serialVersionUID = -2568203115024599915L;

    public PtgLT() {
        ptgId = 0x9;
        record = new byte[1];
        record[0] = 0x9;
    }

    public boolean getIsOperator() {
        return true;
    }

    public boolean getIsPrimitiveOperator() {
        return true;
    }

    public boolean getIsBinaryOperator() {
        return true;
    }

    /**
     * return the human-readable String representation of
     */
    public String getString() {
        return "<";
    }

    public int getLength() {
        return PTG_LT_LENGTH;
    }

    /*  Operator specific calculate method, this one determines if the second-to-top
       operand is less than the top operand;  Returns a PtgBool

   */
    public Ptg calculatePtg(Ptg[] form) {
        try {
            // 20090202 KSC: Handle array formulas
            Object[] o = getValuesFromPtgs(form);
            if (o == null) return new PtgErr(PtgErr.ERROR_VALUE); // some error in value(s)
            if (!o[0].getClass().isArray()) {
                //double[] dub = super.getValuesFromPtgs(form);
                // there should always be only two ptg's in this, error if not.
                if (o.length != 2) {
                    Logger.logWarn("calculating formula failed, wrong number of values in PtgLT");
                    return new PtgErr(PtgErr.ERROR_VALUE);    // 20081203 KSC: handle error's ala Excel return null;
                }
                // blank handling:
                // determine if any of the operands are double - if true,
                // then blank comparisons will be treated as 0's
                boolean isDouble = false;
                for (int i = 0; i < 2 && !isDouble; i++) {
                    if (!form[i].isBlank())
                        isDouble = ((o[i] instanceof Double));
                }
                for (int i = 0; i < 2; i++) {
                    if (form[i].isBlank()) {
                        if (isDouble)
                            o[i] = new Double(0.0);
                        else
                            o[i] = ""; // in this case, empty cells are handled as blank, not zero
                    }
                }

                boolean res;
                //if (dub[0].doubleValue() < dub[1].doubleValue()){
                if (o[0] instanceof Double && o[1] instanceof Double) {
                    res = ((Double) o[0]).doubleValue() < ((Double) o[1]).doubleValue();
                } else {        // string comparison??
                    // This is what Excel does ...
                    if (Formula.isErrorValue(o[0].toString()))
                        return new PtgErr(PtgErr.convertStringToLookupByte(o[0].toString()));
                    if (Formula.isErrorValue(o[1].toString()))
                        return new PtgErr(PtgErr.convertStringToLookupByte(o[1].toString()));
                    // KSC: ExcelTools.transformStringToIntVals does not work in all cases- think of date strings ...
                    res = (o[0].toString().compareTo(o[1].toString()) < 0);
/* KSC: ExcelTools.transformStringToIntVals does not work in all cases- think of date strings ...						
					int[] i1 = ExcelTools.transformStringToIntVals(o[0].toString());
					int[] i2 = ExcelTools.transformStringToIntVals(o[1].toString());
					try {
						res= true;
						for(int k=0;k<i1.length && res;k++){
							res= (i1[k] < i2[k]);
						}
					} catch (ArrayIndexOutOfBoundsException e) {
						res= false;
					}*/
                }

                PtgBool pboo = new PtgBool(res);
                return pboo;
            } else {    // handle array fomulas
                boolean res = false;
                String retArry = "";
                int nArrays = java.lang.reflect.Array.getLength(o);
                if (nArrays != 2) return new PtgErr(PtgErr.ERROR_VALUE);
                int nVals = java.lang.reflect.Array.getLength(o[0]);    // use first array element to determine length of values as subsequent vals might not be arrays
                for (int i = 0; i < nArrays - 1; i += 2) {
                    res = false;
                    Object secondOp = null;
                    boolean comparitorIsArray = o[i + 1].getClass().isArray();
                    if (!comparitorIsArray) secondOp = o[i + 1];
                    for (int j = 0; j < nVals; j++) {
                        Object firstOp = Array.get(o[i], j);    // first array index j
                        if (comparitorIsArray)
                            secondOp = Array.get(o[i + 1], j);    // second array index j

                        if (firstOp instanceof Double && secondOp instanceof Double)
                            res = ((Double) firstOp).compareTo((Double) secondOp) < 0;
                        else { // string comparison?
                            // This is what Excel does ...
                            if (Formula.isErrorValue(o[0].toString()))
                                return new PtgErr(PtgErr.convertStringToLookupByte(o[0].toString()));
                            if (Formula.isErrorValue(o[1].toString()))
                                return new PtgErr(PtgErr.convertStringToLookupByte(o[1].toString()));
                            // KSC: ExcelTools.transformStringToIntVals does not work in all cases- think of date strings ...
                            res = (o[0].toString().compareTo(o[1].toString()) < 0);
		/* KSC: ExcelTools.transformStringToIntVals does not work in all cases- think of date strings ...						
							int[] i1 = ExcelTools.transformStringToIntVals(o[0].toString());
							int[] i2 = ExcelTools.transformStringToIntVals(o[1].toString());
							try {
								res= true;
								for(int k=0;k<i1.length && res;k++){
									res= (i1[k] < i2[k]);
								}
							} catch (ArrayIndexOutOfBoundsException e) {
								res= false;
							}*/
                        }
                        retArry = retArry + res + ",";
                    }
                }
                retArry = "{" + retArry.substring(0, retArry.length() - 1) + "}";
                PtgArray pa = new PtgArray();
                pa.setVal(retArry);
                return pa;
            }
		/*}catch(NumberFormatException e){ 20090203 KSC: Handled above
			try {				
				// Unfortuately <, >, and <> can all deal with strings as well...
				String[] s = getStringValuesFromPtgs(form);
				if (s[0].equalsIgnoreCase(s[1])) return new PtgBool(false);
				int[] i1 = ExcelTools.transformStringToIntVals(s[0]);
				int[] i2 = ExcelTools.transformStringToIntVals(s[1]);
				for(int i=0;i<s.length;i++){
					if (i1[i] > i2[i])return new PtgBool(false);
					if (i1[i] < i2[i])return new PtgBool(true);
				}
				return new PtgBool(true);
		*/
        } catch (Exception ex) {
            return new PtgErr(PtgErr.ERROR_VALUE);
        }
        /*}*/
    }


}