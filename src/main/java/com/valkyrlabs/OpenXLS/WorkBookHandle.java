/*
 * --------- BEGIN COPYRIGHT NOTICE ---------
 * Copyright 2002-2012 Extentech Inc.
 * Copyright 2013 Infoteria America Corp.
 *
 * This file is part of OpenXLS.
 *
 * OpenXLS is free software: you can redistribute it and/or
 * modify
 * it under the terms of the GNU Lesser General Public
 * License as
 * published by the Free Software Foundation, either version
 * 3 of
 * the License, or (at your option) any later version.
 *
 * OpenXLS is distributed in the hope that it will be
 * useful,
 * but WITHOUT ANY WARRANTY; without even the implied
 * warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See
 * the
 * GNU Lesser General Public License for more details.
 *
 * You should have received a copy of the GNU Lesser General
 * Public
 * License along with OpenXLS. If not, see
 * <http://www.gnu.org/licenses/>.
 * ---------- END COPYRIGHT NOTICE ----------
 */
package com.valkyrlabs.OpenXLS;

import java.io.BufferedReader;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import java.io.InputStream;
import java.io.ObjectInputStream;
import java.io.OutputStream;
import java.io.Writer;
import java.net.URL;
import java.nio.ByteBuffer;
import java.nio.ByteOrder;
import java.util.AbstractList;
import java.util.Arrays;
import java.util.Hashtable;
import java.util.Iterator;
import java.util.List;
import java.util.Vector;

import com.valkyrlabs.formats.LEO.BlockByteReader;
import com.valkyrlabs.formats.LEO.InvalidFileException;
import com.valkyrlabs.formats.LEO.LEOFile;
import com.valkyrlabs.formats.XLS.BiffRec;
import com.valkyrlabs.formats.XLS.BookProtectionManager;
import com.valkyrlabs.formats.XLS.Boundsheet;
import com.valkyrlabs.formats.XLS.Condfmt;
import com.valkyrlabs.formats.XLS.Font;
import com.valkyrlabs.formats.XLS.Formula;
import com.valkyrlabs.formats.XLS.Hlink;
import com.valkyrlabs.formats.XLS.Mergedcells;
import com.valkyrlabs.formats.XLS.Mulblank;
import com.valkyrlabs.formats.XLS.Name;
import com.valkyrlabs.formats.XLS.OOXMLAdapter;
import com.valkyrlabs.formats.XLS.OOXMLReader;
import com.valkyrlabs.formats.XLS.OOXMLWriter;
import com.valkyrlabs.formats.XLS.PivotCache;
import com.valkyrlabs.formats.XLS.Sxview;
import com.valkyrlabs.formats.XLS.WorkBookFactory;
import com.valkyrlabs.formats.XLS.XLSConstants;
import com.valkyrlabs.formats.XLS.Xf;
import com.valkyrlabs.formats.XLS.charts.Chart;
import com.valkyrlabs.formats.XLS.charts.OOXMLChart;
import com.valkyrlabs.toolkit.JFileWriter;
import com.valkyrlabs.toolkit.Logger;
import com.valkyrlabs.toolkit.ProgressListener;
import com.valkyrlabs.toolkit.ResourceLoader;
import com.valkyrlabs.toolkit.StringTool;
import com.valkyrlabs.toolkit.TempFileManager;

/**
 * The WorkBookHandle provides a handle to the XLS file and includes convenience
 * methods for working with the WorkSheets and Cell values within the XLS file.
 * For example: <br>
 * <code>
 * WorkBookHandle  book  = new WorkBookHandle("testxls.xls");<br>
 * WorkSheetHandle sheet = book.getWorkSheet("Sheet1");<br>
 * CellHandle      cell  = sheet.getCell("B22");<br>
 * </code>
 *
 * <p>
 * By default, OpenXLS will lock open WorkBook files. To close the file after
 * parsing and work with a temporary file instead, use the following setting:
 * <br>
 * <code>
 * System.getProperties().put(WorkBookHandle.USETEMPFILE, "true");
 * </code><br>
 * If you enable this mode you will need to periodically clean up the generated
 * temporary files in your working directory. All OpenXLS temporary file names
 * begin with "ExtenXLS_".
 */
public class WorkBookHandle extends DocumentHandle implements WorkBook, Handle {

    /**
     * Format constant for BIFF8 (Excel '97-2007).
     */
    public static final int FORMAT_XLS = 100;
    /**
     * Format constant for normal OOXML (Excel 2007).
     */
    public static final int FORMAT_XLSX = 101;
    /**
     * Format constant for macro-enabled OOXML (Excel 2007).
     */
    public static final int FORMAT_XLSM = 102;
    /**
     * Format constant for OOXML template (Excel 2007).
     */
    public static final int FORMAT_XLTX = 103;
    /**
     * Format constant for macro-enabled OOXML template (Excel 2007).
     */
    public static final int FORMAT_XLTM = 104;
    /**
     * A Writer to which a record dump should be written on input. This is used by
     * the dumping code in WorkBookFactory.
     */
    public static Writer dump_input = null;
    public static java.text.SimpleDateFormat simpledateformat = new java.text.SimpleDateFormat(); // static
    /**
     * How many recursion levels to allow formulas to be calculated before throwing
     * a circular reference error
     */
    public static int RECURSION_LEVELS_ALLOWED = 107;
    /**
     * This is
     **/
    public static String CONVERTMULBLANKS = "deprecated";
    private static byte[] protobook;
    private static byte[] protochart;
    private static byte[] protosheet;
    protected com.valkyrlabs.formats.XLS.WorkBook mybook;
    protected LEOFile myLEOFile;
    // to
    // reuse
    protected WorkBookFactory myfactory = null;
    protected ProgressListener plist;
    Hashtable<String, WorkSheetHandle> sheethandles = new Hashtable<String, WorkSheetHandle>();

    /**
     * Default constructor creates a new, empty Spreadsheet with
     * <p>
     * 3 WorkSheets: "Sheet1","Sheet2",and "Sheet3".
     */
    public WorkBookHandle() {
        // Xf.DEFAULTIXFE= 15; // reset to default in cases of
        // having previously read
        // Excel2007 template which may have set defaultXF
        // differently
        this.initDefault();
    }

    /**
     * Constructor creates a new, empty Spreadsheet with 3 worksheets: "Sheet1",
     * "Sheet2" and "Sheet3" <br>
     * This version allows flagging the workbook as Excel 2007 format. <br>
     * Excel 2007 format contains larger maximum column and row contraints, for
     * example. <br>
     * Even if the workbook is set to Excel 2007 format, it is still possible to
     * write out the file as a BIFF8 (Excel 97-2003) file, but unsupported features
     * will be dropped, and some files could experience corruption.
     *
     * @param boolean Excel2007 - true if set to Excel 2007 version
     */
    public WorkBookHandle(boolean Excel2007) {
        this.initDefault();
        this.setIsExcel2007(Excel2007);
    }

    /**
     * constructor which takes an InputStream containing the bytes of a valid XLS
     * file.
     *
     * @param InputStream contains the valid BIFF8 bytes for reading
     */
    public WorkBookHandle(InputStream inx) {
        this.initFromStream(inx);
    }

    /**
     * Create a new WorkBookHandle from the byte array passed in. Byte array passed
     * in must contain a valid xls or xlsx workbook file
     *
     * @param byte[] byte array containing the valid XLS or XLSX file for reading
     */
    public WorkBookHandle(byte[] barray) {
        initializeFromByteArray(barray);
    }

    /**
     * Fetches a workbook from a URL
     * <p>
     * If you need to authenticate your URL connection first then use the
     * InputStream constructor
     *
     * @param urlx
     * @return
     * @throws Exception
     */
    public WorkBookHandle(URL url) {
        /*
         * OK, both this method and the (inputstream) constructor
         * set a temp file, is
         * this not possible to do without hitting the disk? TODO:
         * look into fix
         */
        this(getFileFromURL(url));
    }

    /**
     * Constructor which takes the XLS file name(
     *
     * @param String filePath the name of the XLS file to read
     */
    public WorkBookHandle(String filePath) {
        this(filePath, 0);
    }

    /**
     * constructor which takes the XLS file name and has an optional debug setting
     * to assist with output. Setting this value will cause verbose logging and is
     * discouraged unless required for support.
     *
     * @param String filePath the name of the XLS file to read
     * @param Debug  level
     */
    public WorkBookHandle(String filePath, int debug) {
        this.setDebugLevel(debug);
        File f = new File(filePath);
        this.initFromFile(f);
        this.file = f; // XXX KSC: Save for potential re-input of pass-through
        // ooxml files

    }

    /**
     * constructor which takes the XLS file
     *
     * @param File  the XLS file to read
     * @param Debug level
     */
    public WorkBookHandle(File fx) {
        this.initFromFile(fx);
    }

    /**
     * Constructor which takes a ProgressListener which monitors the progress of
     * creating a new Excel file.
     *
     * @param ProgressListener object which is monitoring progress of WorkBook read
     */
    public WorkBookHandle(ProgressListener pn) {
        this.plist = pn;
        try {
            byte[] b = getPrototypeBook();
            ByteBuffer bbf = ByteBuffer.wrap(b);
            bbf.order(ByteOrder.LITTLE_ENDIAN);
            myLEOFile = new LEOFile(bbf);
        } catch (Exception e) {
            throw new InvalidFileException(
                    "WorkBook could not be instantiated: " + e.toString());
        }
        this.initFromLeoFile(myLEOFile);
    }

    /**
     * Constructor which takes the XLS file name
     * <p>
     * and a ProgressListener which monitors the progress of reading the Excel file.
     *
     * @param String           fname the name of the XLS file to read
     * @param ProgressListener object which is monitoring progress of WorkBook read
     */
    public WorkBookHandle(String fname, ProgressListener pn) {
        this.plist = pn;
        this.initFromFile(new File(fname));
    }

    /**
     * For internal creation of a workbook handle from
     *
     * @param leo
     */
    protected WorkBookHandle(LEOFile leo) {
        this.initFromLeoFile(leo);
    }

    protected static byte[] getPrototypeBook() throws IOException {
        if (protobook == null)

            protobook = ResourceLoader
                    .getBytesFromJar("/com/valkyrlabs/OpenXLS/templates/prototysspe.ser");

        return protobook;
    }

    protected static byte[] getPrototypeSheet() {
        if (protosheet == null)
            try {
                WorkBookHandle bookhandle = new WorkBookHandle();
                com.valkyrlabs.formats.XLS.WorkBook book = bookhandle.getWorkBook();
                Boundsheet sheet = book.getWorkSheetByNumber(0);
                protosheet = sheet.getSheetBytes();
            } catch (Exception e) {
                throw new RuntimeException(e);
            }

        return protosheet;
    }

    protected static byte[] getPrototypeChart() {
        if (protochart == null) {
            try {
                byte[] bookbytes = ResourceLoader
                        .getBytesFromJar("/com/valkyrlabs/OpenXLS/templates/prototypechart.ser");
                WorkBookHandle chartBook = new WorkBookHandle(bookbytes);
                ChartHandle ch = chartBook.getCharts()[0];
                protochart = ch.getSerialBytes();
                return protochart;
            } catch (IOException e) {
                Logger.logErr("Unable to get default chart bytes");
            }
        }
        return protochart;
    }

    /**
     * Set the recursion levels allowed for formulas calculated in this workbook
     * before a circular reference is thrown.
     * <p>
     * Default setting is 250 levels of recursion
     *
     * @param recursion_allowed
     */
    public static void setFormulaRecursionLevels(int recursion_allowed) {
        RECURSION_LEVELS_ALLOWED = recursion_allowed;
    }

    /**
     * Gets the internal Factory object.
     * <p>
     * <strong>WARNING:</strong> This method is not part of the public API. Its use
     * is not supported and behavior is subject to change.
     * </p>
     */
    public WorkBookFactory getFactory() {
        return myfactory;
    }

    /**
     * Gets the internal LEOFile object.
     * <p>
     * <strong>WARNING:</strong> This method is not part of the public API. Its use
     * is not supported and behavior is subject to change.
     * </p>
     */
    public LEOFile getLEOFile() {
        return myLEOFile;
    }

    /**
     * Searches all Cells in the workbook for the string occurrence and replaces
     * with the replacement text.
     *
     * @return the number of replacements that were made
     */
    public int searchAndReplace(String searchfor, String replacewith) {
        CellHandle[] cx = getCells();
        int foundcount = 0;
        for (int t = 0; t < cx.length; t++) {
            if (!(cx[t].getCell() instanceof Formula)) {
                // find the string
                if (!cx[t].isNumber()) {
                    String v = cx[t].getStringVal();
                    if (v.indexOf(searchfor) > -1) {
                        cx[t].setVal(StringTool
                                .replaceText(v, searchfor, replacewith));
                        foundcount++;
                    }
                }
            }
        }
        return foundcount;
    }

    /**
     * Returns all strings that are in the SharedStringTable for this workbook. The
     * SST contains all standard string records in cells, but may not include such
     * things as strings that are contained within formulas. This is useful for such
     * things as full text indexing of workbooks
     *
     * @return Strings in the workbook.
     */
    public String[] getAllStrings() {
        return this.mybook.getAllStrings();
    }

    /**
     * Get either the default color table, or the color table from the custom
     * palatte(if exists) from the WorkBook
     *
     * @return
     */
    @Override
    public java.awt.Color[] getColorTable() {
        return this.getWorkBook().getColorTable();
    }

    /**
     * Returns whether this book uses the 1904 date format.
     *
     * @deprecated Use {@link #getDateFormat()} instead.
     */
    @Deprecated
    public boolean is1904() {
        return mybook.is1904();
    }

    /**
     * Gets the date format used by this book.
     */
    @Override
    public DateConverter.DateFormat getDateFormat() {
        return mybook.getDateFormat();
    }

    /**
     * Returns the lowest version of Excel compatible with the input file.
     *
     * @return an Excel version string
     */
    public String getXLSVersionString() {
        return mybook.getXLSVersionString();
    }

    /**
     * Return useful statistics about this workbook.
     *
     * @param use html line breaks
     * @return a string contatining various statistics.
     */
    public String getStats(boolean usehtml) {
        return mybook.getStats(usehtml);
    }

    /**
     * Return useful statistics about this workbook.
     *
     * @return a string contatining various statistics.
     */
    public String getStats() {
        return mybook.getStats();
    }

    /**
     * Returns the Cell at the specified Location
     *
     * @param address
     * @return
     * @throws SheetNotFoundException
     */
    @Override
    public CellHandle getCell(String address) throws CellNotFoundException, SheetNotFoundException {
        int shtpos = address.indexOf("!");
        if (shtpos < 0)
            throw new CellNotFoundException(address
                    + " not found.  You need to specify a location in the format: Sheet1!A1");
        String sheetstr = address.substring(0, shtpos);
        WorkSheetHandle sht = this.getWorkSheet(sheetstr);
        String celstr = address.substring(shtpos + 1);
        return sht.getCell(celstr);
    }

    /**
     * Returns an Array of the CellRanges existing in this WorkBook specifically the
     * Ranges referenced in Formulas, Charts, and Named Ranges.
     * <p>
     * This is necessary to allow for automatic updating of references when
     * adding/removing/moving Cells within these ranges, as well as shifting
     * references to Cells in Formulas when Formula records are moved.
     *
     * @return all existing Cell Range references used in Formulas, Charts, and
     *         Names
     */
    public CellRange[] getCellRanges() {
        return this.mybook.getRefTracker().getCellRanges();
    }

    /**
     * get a handle to a PivotTable in the WorkBook
     *
     * @param String name of the PivotTable
     * @return PivotTable the PivotTable
     */
    @Override
    public PivotTableHandle getPivotTable(String ptname) throws PivotTableNotFoundException {
        Sxview st = mybook.getPivotTableView(ptname);
        if (st == null)
            throw new PivotTableNotFoundException(ptname);
        return new PivotTableHandle(st, this);
    }

    /**
     * get an array of handles to all PivotTables in the WorkBook
     *
     * @return PivotTable[] all of the WorkBooks PivotTables
     */
    @Override
    public PivotTableHandle[] getPivotTables() throws PivotTableNotFoundException {
        Sxview[] sxv = mybook.getAllPivotTableViews();
        if (sxv == null || sxv.length == 0)
            throw new PivotTableNotFoundException(
                    "There are no PivotTables defined in: " + this.getName());
        PivotTableHandle[] pth = new PivotTableHandle[sxv.length];
        for (int t = 0; t < pth.length; t++) {
            pth[t] = new PivotTableHandle(sxv[t], this);
        }
        return pth;
    }

    /**
     * Get the calculation mode for the workbook.
     * <p>
     * CALCULATE_ALWAYS is the default for new workbooks. Calling Cell.getVal() will
     * calculate formulas if they exist within the cell.
     * <p>
     * CALCULATE_EXPLICIT will return present value of the cell. Formula calculation
     * will only occur when explicitly called through the Formula Handle
     * <p>
     * WorkBookHandle.CALCULATE_ALWAYS -- recalc every time the cell value is
     * requested (no cacheing) WorkBookHandle.CALCULATE_EXPLICIT -- recalc only when
     * FormulaHandle.calculate() called WorkBookHandle.CALCULATE_AUTO -- only recac
     * when changes
     *
     * @param CalcMode Calculation mode to use in workbook.
     */
    @Override
    public int getFormulaCalculationMode() {
        return mybook.getCalcMode();
    }

    /**
     * Set the calculation mode for the workbook.
     * <p>
     * CALCULATE_AUTO is the default for new workbooks. Calling Cell.getVal() will
     * calculate formulas if they exist within the cell.
     * <p>
     * CALCULATE_EXPLICIT will return present, cached value of the cell. Formula
     * calculation will ONLY occur when explicitly called through the Formula
     * Handle.calculate() method.
     * <p>
     * CALCULATE_ALWAYS will ignore the cache and force a recalc every time a cell
     * value is requested.
     * <p>
     * <p>
     * WorkBookHandle.CALCULATE_AUTO WorkBookHandle.CALCULATE_ALWAYS
     * WorkBookHandle.CALCULATE_EXPLICIT
     *
     * @param CalcMode Calculation mode to use in workbook.
     */
    @Override
    public void setFormulaCalculationMode(int CalcMode) {
        mybook.setCalcMode(CalcMode);
    }

    /**
     * set the workbook to protected mode
     * <p>
     * Note: the password cannot be decrypted or changed in Excel -- protection can
     * only be set/removed using OpenXLS
     *
     * @param boolean whether to protect the book
     */
    @Override
    public void setProtected(boolean protect) {
        // TODO: Check that this behavior is correct
        // This is what the old implementation did

        BookProtectionManager protector = mybook.getProtectionManager();

        // Excel default... no kidding!
        if (protect)
            protector.setPassword("VelvetSweatshop");
        else
            protector.setPassword(null);

        protector.setProtected(protect);
    }

    /**
     * set Default row height in twips (=1/20 of a point)
     * <p>
     * Note: only affects undefined Rows containing Cells
     *
     * @param int Default Row Height
     */
    // should be a double as Excel units are 1/20 of what is
    // stored in
    // defaultrowheight
    // e.g. 12.75 is Excel Units, twips = 12.75*20 = 256
    // (approx)
    // should expect users to use Excel units and target method
    // do the 20*
    // conversion
    @Override
    public void setDefaultRowHeight(int t) {
        mybook.setDefaultRowHeight(t);
    }

    /**
     * Set the default column width across all worksheets <br/>
     * This setting is a worksheet level setting, so will be applied to all existing
     * worksheets. individual worksheets can also be set using
     * WorkSheetHandle.setDefaultColWidth
     * <p>
     * This setting is roughly the width of the character '0' The default width of a
     * column is 8.
     */
    @Override
    public void setDefaultColWidth(int t) {
        mybook.setDefaultColWidth(t);
    }

    /**
     * Returns a Formula Handle
     *
     * @return FormulaHandle a formula handle in the WorkBook
     */
    public FormulaHandle getFormulaHandle(String celladdress) throws FormulaNotFoundException {
        Formula formula = mybook.getFormula(celladdress);
        return new FormulaHandle(formula, this);
    }

    /**
     * Returns all ImageHandles in the workbook
     *
     * @return
     */
    public ImageHandle[] getImages() {
        List<ImageHandle> ret = new Vector<ImageHandle>();
        for (int t = 0; t < this.getNumWorkSheets(); t++) {
            try {
                ImageHandle[] r = this.getWorkSheet(t).getImages();
                for (int x = 0; x < r.length; x++)
                    ret.add(r[x]);
            } catch (Exception ex) {
            }
        }
        ImageHandle[] retx = new ImageHandle[ret.size()];
        ret.toArray(retx);
        return retx;
    }

    /**
     * Returns an ImageHandle for manipulating images in the WorkBook
     *
     * @param imagename
     * @return
     */
    public ImageHandle getImage(String imagename) throws ImageNotFoundException {
        for (int t = 0; t < this.getNumWorkSheets(); t++) {
            try {
                ImageHandle[] r = this.getWorkSheet(t).getImages();
                for (int x = 0; x < r.length; x++)
                    if (r[x].getName().equals(imagename))
                        return r[x];
            } catch (Exception ex) {
            }
        }
        throw new ImageNotFoundException(
                "Image not found: " + imagename + " in " + this.toString());
    }

    /**
     * Returns a Named Range Handle
     *
     * @return NameHandle a Named range in the WorkBook
     */
    @Override
    public NameHandle getNamedRange(String rangename) throws CellNotFoundException {
        Name nand = mybook.getName(rangename.toUpperCase()); // case-insensitive
        if (nand == null)
            throw new CellNotFoundException(rangename);
        return new NameHandle(nand, this);
    }

    /**
     * Returns a Named Range Handle if it exists in the specified scope.
     * <p>
     * This can be used to distinguish between multiple named ranges with the same
     * name but differing scopes
     *
     * @return NameHandle a Named range in the WorkBook that exists in the scope
     */
    public NameHandle getNamedRangeInScope(String rangename) throws CellNotFoundException {
        Name nand = mybook.getScopedName(rangename);
        if (nand == null)
            throw new CellNotFoundException(rangename);
        return new NameHandle(nand, this);
    }

    /**
     * Create a named range in the workbook
     * <p>
     * Note that the named range designation can conform to excel specs, that is,
     * boolean values, references, or string variables can be set. Remember to
     * utilize the sheet name when setting referential names.
     * <p>
     * <p>
     * NameHandle nh = createNamedRange("cellRange", "Sheet1!A1:B3"); NameHandle nh
     * = createNamedRange("trueRange", "=true");
     *
     * @param name     The name that should be used to reference this named range
     * @param rangeDef Range of the cells for this named range, in excel syntax
     *                 including
     *                 sheet name, ie "Sheet1!A1:D1"
     * @return NameHandle for modifying the named range
     */
    public NameHandle createNamedRange(String name, String rangeDef) {
        NameHandle nh = new NameHandle(name, rangeDef, this);
        return nh;
    }

    /**
     * Returns a Chart Handle
     *
     * @return ChartHandle a Chart in the WorkBook
     */
    // KSC: NOTE: this methodology needs work as a book may
    // contain charts in
    // different sheets containing the same name
    // TODO: rethink
    @Override
    public ChartHandle getChart(String chartname) throws ChartNotFoundException {
        return new ChartHandle(mybook.getChart(chartname), this);
    }

    /**
     * Returns all Chart Handles contained in the WorkBook
     *
     * @return ChartHandle[] an array of all Charts in the WorkBook
     */
    @Override
    public ChartHandle[] getCharts() {
        AbstractList<?> cv = mybook.getChartVect();
        ChartHandle[] cht = new ChartHandle[cv.size()];
        for (int x = 0; x < cv.size(); x++) {
            cht[x] = new ChartHandle((Chart) cv.get(x), this);
        }
        return cht;
    }

    /**
     * retrieve a ChartHandle via id
     *
     * @param id
     * @return
     * @throws ChartNotFoundException
     */
    public ChartHandle getChartById(int id) throws ChartNotFoundException {
        AbstractList<?> cv = mybook.getChartVect();
        Chart cht = null;
        for (int x = 0; x < cv.size(); x++) {
            cht = (Chart) cv.get(x);
            if (cht.getId() == id)
                return new ChartHandle(cht, this);
        }
        throw new ChartNotFoundException("Id " + id);
    }

    /**
     * Returns all Named Range Handles
     *
     * @return NameHandle[] all of the Named ranges in the WorkBook
     */
    @Override
    public NameHandle[] getNamedRanges() {
        Name[] nand = mybook.getNames();
        NameHandle[] nands = new NameHandle[nand.length];
        for (int x = 0; x < nand.length; x++) {
            nands[x] = new NameHandle(nand[x], this);
        }
        return nands;
    }

    /**
     * Returns all Named Range Handles scoped to WorkBook.
     * <p>
     * Note this will not include worksheet scoped named ranges
     *
     * @return NameHandle[] all of the Named ranges that are scoped to WorkBook
     */
    public NameHandle[] getNamedRangesInScope() {
        Name[] nand = mybook.getWorkbookScopedNames();
        NameHandle[] nands = new NameHandle[nand.length];
        for (int x = 0; x < nand.length; x++) {
            nands[x] = new NameHandle(nand[x], this);
        }
        return nands;
    }

    /**
     * Returns the name of this WorkBook
     *
     * @return String name of WorkBook
     */
    @Override
    public String getName() {
        if (name != null)
            return name;
        return "New Spreadsheet";
    }

    /**
     * Returns an array containing all cells in the WorkBook
     *
     * @return CellHandle array of all book cells
     */
    @Override
    public CellHandle[] getCells() {
        BiffRec[] allcz = this.mybook.getCells();
        CellHandle[] ret = new CellHandle[allcz.length];
        Mulblank aMul = null;
        short c = -1;
        for (int t = 0; t < ret.length; t++) {
            ret[t] = new CellHandle(allcz[t], this);
            if (allcz[t].getOpcode() == XLSConstants.MULBLANK) {
                // handle Mulblanks: ref a range of cells; to get correct
                // cell address,
                // traverse thru range and set cellhandle ref to correct
                // column
                if (allcz[t] == aMul) {
                    c++;
                } else {
                    aMul = (Mulblank) allcz[t];
                    c = (short) aMul.getColFirst();
                }
                ret[t].setBlankRef(c); // for Mulblank use only -sets correct
                // column reference for multiple blank
                // cells
                // ...
            }
        }
        return ret;
    }

    /**
     * Returns the number of Cells in this WorkBook
     *
     * @return int number of Cells
     */
    @Override
    public int getNumCells() {
        return mybook.getNumCells();
    }

    /**
     * Returns whether the sheet selection tabs should be shown.
     */
    public boolean showSheetTabs() {
        return mybook.showSheetTabs();
    }

    /**
     * Sets whether the sheet selection tabs should be shown.
     */
    public void setShowSheetTabs(boolean show) {
        mybook.setShowSheetTabs(show);
    }

    /**
     * Gets the spreadsheet as a byte array in BIFF8 (Excel '97-2003) format.
     *
     * @deprecated Writing the spreadsheet to a byte array uses a great deal of
     *             memory and generally provides no benefit over streaming output.
     *             Use the {@link #write} family of methods instead. If you need a
     *             byte array use {@link ByteArrayOutputStream}.
     */
    @Override
    @Deprecated
    public byte[] getBytes() {
        try {
            ByteArrayOutputStream bout = new ByteArrayOutputStream();
            writeBytes(bout);

            return bout.toByteArray();
        } catch (Exception e1) {
            Logger.logErr("Getting Spreadsheet bytes failed.", e1);
            return null;
        }
    }

    /**
     * Writes the document to the given path. If the filename ends with ".xlsx" or
     * ".xlsm", the workbook will be written as OOXML (XLSX). Otherwise it will be
     * written as BIFF8 (XLS). For OOXML, if the file has a VBA project the file
     * extension must be ".xlsm". It will be changed if necessary.
     *
     * @param path the path to which the document should be written
     * @deprecated The filename-based format choosing is counter-intuitive and
     *             failure-prone. Use {@link #write(OutputStream, int)} instead.
     */
    @Deprecated
    public void write(String path) {
        String ext = path.toLowerCase();
        write(path, ext.endsWith(".xlsx") || ext.endsWith(".xlsm"));
    }

    /**
     * Writes the document to the given file in either XLS or XLSX. For OOXML, if
     * the file has a VBA project the file extension must be ".xlsm". It will be
     * changed if necessary.
     *
     * @param path  the path to which the document should be written
     * @param ooxml If <code>true</code>, write as OOXML (XLSX). Otherwise, write as
     *              BIFF8 (XLS).
     * @deprecated The boolean format parameter is not flexible enough to represent
     *             all supported formats. Use {@link #write(File, int)} instead.
     */
    @Deprecated
    public void write(String path, boolean ooxml) {
        int format;
        if (ooxml) {
            if (this.getIsExcel2007())
                format = this.getFormat(path);
            else
                format = FORMAT_XLSX;

            if (!OOXMLAdapter.hasMacros(this))
                path = StringTool.replaceExtension(path, ".xlsx");
            else // it's a macro-enabled workbook
                path = StringTool.replaceExtension(path, ".xlsm");
        } else
            format = FORMAT_XLS;

        try {
            this.write(new File(path), format);
        } catch (Exception e) {
            throw new WorkBookException("error writing workbook",
                    WorkBookException.WRITING_ERROR, e);
        }
    }

    /**
     * Writes the document to the given stream in either XLS or XLSX format.
     *
     * @param dest  the stream to which the document should be written
     * @param ooxml If <code>true</code>, write as OOXML (XLSX). Otherwise, write as
     *              BIFF8 (XLS).
     * @deprecated The boolean format parameter is not flexible enough to represent
     *             all supported formats. Use {@link #write(OutputStream, int)}
     *             instead.
     */
    @Deprecated
    public void write(OutputStream dest, boolean ooxml) {
        int format;
        if (ooxml) {
            if (this.getIsExcel2007())
                format = this.getFormat();
            else
                format = FORMAT_XLSX;
        } else
            format = FORMAT_XLS;

        try {
            if (format > WorkBookHandle.FORMAT_XLS && this.file != null) {
                OOXMLAdapter.refreshPassThroughFiles(this);
            }
            this.write(dest, format);
        } catch (Exception e) {
            throw new WorkBookException("error writing workbook",
                    WorkBookException.WRITING_ERROR, e);
        }
    }

    /**
     * Gets the constant representing this document's native format.
     */
    @Override
    public int getFormat() {
        String name = this.getFileName().toLowerCase();

        if (this.getIsExcel2007()) {
            if (OOXMLAdapter.hasMacros(this))
                return name.endsWith(".xltm") ? FORMAT_XLTM : FORMAT_XLSM;
            else
                return name.endsWith(".xltx") ? FORMAT_XLTX : FORMAT_XLSX;
        } else
            return FORMAT_XLS;
    }

    /**
     * Gets the constant representing this document's desired format
     */
    public int getFormat(String path) {
        if (path == null)
            return getFormat();
        if (this.getIsExcel2007()) {
            if (OOXMLAdapter.hasMacros(this))
                return path.endsWith(".xltm") ? FORMAT_XLTM : FORMAT_XLSM;
            else
                return path.endsWith(".xltx") ? FORMAT_XLTX : FORMAT_XLSX;
        } else
            return FORMAT_XLS;
    }

    @Override
    public String getFileExtension() {
        switch (this.getFormat()) {
            case FORMAT_XLSX:
                return ".xlsx";
            case FORMAT_XLSM:
                return ".xlsm";
            case FORMAT_XLTX:
                return ".xltx";
            case FORMAT_XLTM:
                return ".xltm";
            case FORMAT_XLS:
                return ".xls";
            default:
                return "";
        }
    }

    /**
     * Writes the document to the given stream in the requested format.
     * <p>
     * format choices:
     * <p>
     * WorkBookHandle.FORMAT_XLS for 2003 and previous versions <br>
     * WorkBookHandle.FORMAT_XLSX for non-macro-enabled 2007 version <br>
     * WorkBookHandle.FORMAT_XLSM for macro-enabled 2007 version <br>
     * WorkBookHandle.FORMAT_XLTM for macro-enabled 2007 templates. <br>
     * WorkBookHandle.FORMAT_XLTX for 2007 templates,
     * <p>
     * <b>IMPORTANT NOTE:</b> if the resulting filename contains the .XLSM extension
     * <br>
     * the WorkBook <b>MUST</b> be written in FORMAT_XLSM; otherwise open errors
     * will occur
     * <p>
     * <b>NOTE:</b> If the format is FORMAT_XLSX and the filename contains macros
     * <br>
     * the file will be written as Macro-Enabled i.e. in FORMAT_XLSM. In these
     * cases, <br>
     * the filename must contain the .XLSM extension
     *
     * @param dest   the stream to which the document should be written
     * @param format the constant representing the desired output format
     * @throws IllegalArgumentException if the given type code is invalid
     * @throws IOException              if an error occurs while writing to the
     *                                  stream
     */
    @Override
    public void write(OutputStream dest, int format) throws IOException {
        if (format == FORMAT_NATIVE)
            format = this.getFormat();

        switch (format) {
            case FORMAT_XLSX:
            case FORMAT_XLSM:
            case FORMAT_XLTX:
            case FORMAT_XLTM:
                try {
                    if (this.file != null)
                        OOXMLAdapter.refreshPassThroughFiles(this);

                    OOXMLWriter adapter = new OOXMLWriter();
                    adapter.setFormat(format);
                    adapter.getOOXML(this, dest);
                } catch (IOException e) {
                    throw e;
                } catch (Exception e) {
                    // TODO: OOXMLAdapter only throws IOException, change its
                    // throws
                    throw new WorkBookException("error writing workbook",
                            WorkBookException.WRITING_ERROR, e);
                }
                break;

            case FORMAT_XLS:
                try {
                    this.mybook.getStreamer().writeOut(dest);
                } catch (com.valkyrlabs.OpenXLS.WorkBookException e) {
                    Throwable cause = e.getCause();
                    if (cause instanceof IOException)
                        throw (IOException) cause;
                    throw e;
                }
                break;

            default:
                throw new IllegalArgumentException("unknown output format");
        }
    }

    /**
     * Writes the document to the given stream in the requested OOXML format.
     *
     * @param dest   the stream to which the document should be written
     * @param format the constant representing the desired output format
     * @throws IllegalArgumentException if the given type code is invalid
     * @throws IOException              if an error occurs while writing to the
     *                                  stream
     * @deprecated This method is like {@link #write(OutputStream, int)} except it
     *             only supports OOXML formats. Use that instead.
     */
    @Deprecated
    public void writeXLSXBytes(OutputStream dest, int format) throws Exception {
        this.write(dest, format);
    }

    /**
     * Writes the document to the given stream in the default OOXML format.
     *
     * @param dest the stream to which the document should be written
     * @throws IOException if an error occurs while writing to the stream
     * @deprecated Use {@link #write(OutputStream, int}) instead.
     */
    @Deprecated
    public void writeXLSXBytes(OutputStream dest) throws Exception {
        this.write(dest, true);
    }

    /**
     * Returns whether the underlying spreadsheet is in Excel 2007 format by
     * default.
     * <p>
     * Even if this method returns true, it is still possible to write out the file
     * as a BIFF8 (Excel 97-2003) file, but unsupported features will be dropped,
     * and some files could experience corruption.
     *
     * @return whether the underlying spreadsheet is Excel 2007 format
     */
    public boolean getIsExcel2007() {
        return this.mybook.getIsExcel2007();
    }

    /**
     * Sets whether this Workbook is in Excel 2007 format. Excel 2007 format
     * contains larger maximum column and row contraints, for example. <br>
     * Even if the workbook is set to Excel 2007 format, it is still possible to
     * write out the file as a BIFF8 (Excel 97-2003) file, but unsupported features
     * will be dropped, and some files could experience corruption.
     *
     * @param isExcel2007
     */
    public void setIsExcel2007(boolean isExcel2007) {
        this.mybook.setIsExcel2007(isExcel2007);
    }

    /**
     * another handle to the useful ability to load a book from the prorotype bytes
     */
    protected void initDefault() {
        try {
            byte[] b = getPrototypeBook();
            if (b == null) {
                throw new com.valkyrlabs.OpenXLS.WorkBookException(
                        "Unable to load prototype workbook.",
                        WorkBookException.LICENSING_FAILED);
            }
            ByteBuffer bbf = ByteBuffer.wrap(b);
            bbf.order(ByteOrder.LITTLE_ENDIAN);
            myLEOFile = new LEOFile(bbf);
        } catch (Exception e) {
            throw new InvalidFileException(
                    "WorkBook could not be instantiated: " + e.toString());
        }
        this.initFromLeoFile(myLEOFile);
    }

    /**
     * Initialization of this workbook handle from a leoFile;
     */
    private void initFromLeoFile(LEOFile leo) {
        this.myLEOFile = leo;
        try {
            BlockByteReader bar = myLEOFile.getXLSBlockBytes();
            this.initBytes(bar);
            this.setIsExcel2007(false);
            myLEOFile.clearAfterInit();
        } catch (Exception e) {
            if (e instanceof com.valkyrlabs.OpenXLS.WorkBookException)
                throw (com.valkyrlabs.OpenXLS.WorkBookException) e;
            throw new com.valkyrlabs.OpenXLS.WorkBookException(
                    "ERROR: instantiating WorkBookHandle failed: " + e,
                    WorkBookException.UNSPECIFIED_INIT_ERROR, e);
        }
    }

    /**
     * Initialize this workbook from a stream, unfortunately our byte backer
     * requires a file, so create a tempfile and init from that
     */
    protected void initFromStream(InputStream input) {
        try {
            File target = TempFileManager.createTempFile("WBP", ".tmp");

            JFileWriter.writeToFile(input, target);
            this.initFromFile(target.getAbsoluteFile());
            if (this.myLEOFile != null)// it would be if XLSX or XLSM ...
                // 20090323 KSC
                this.myLEOFile.closefb();
            // this.myLEOFile.close(); // close now flushes buffers +
            // storages ...
            input.close();

            File fdel = new File(target.toString());
            if (!fdel.delete()) {
                if (this.DEBUGLEVEL > DEBUG_LOW)
                    Logger.logWarn("Could not delete tempfile: "
                            + target.toString());
            }
        } catch (IOException ex) {
            Logger.logErr("Initializing WorkBookHandle failed.", ex);
        }
    }

    /**
     * Protected method that handles WorkBookHandle(byte[]) constructor
     *
     * @param barray
     */
    protected void initializeFromByteArray(byte[] barray) {
        // check first bytes to see if this is a zipfile (OOXML)
        if ((char) barray[0] == 'P' && (char) barray[1] == 'K') {
            try {
                // added "." fixes Baxter Open Bug [BugTracker 2909]
                File ftmp = TempFileManager.createTempFile("WBP", ".tmp");
                FileOutputStream fous = new FileOutputStream(ftmp);
                fous.write(barray);
                fous.flush();
                fous.close();
                this.initFromFile(ftmp);
                return;
            } catch (Exception e) {
                Logger.logErr("Could not parse XLSX from bytes."
                        + e.toString());
                return;
            }
        }

        ByteBuffer bbf = ByteBuffer.wrap(barray);
        bbf.order(ByteOrder.LITTLE_ENDIAN);
        myLEOFile = new LEOFile(bbf);
        if (myLEOFile.hasWorkBook()) {
            try {
                BlockByteReader bar = myLEOFile.getXLSBlockBytes();
                this.initBytes(bar);
            } catch (Throwable e) {
                if (e instanceof OutOfMemoryError)
                    throw (Error) e;
                if (e instanceof WorkBookException)
                    throw (WorkBookException) e;
                String errstr = "Instantiating WorkBookHandle failed: "
                        + e.toString();
                throw new com.valkyrlabs.OpenXLS.WorkBookException(errstr,
                        WorkBookException.UNSPECIFIED_INIT_ERROR);
            }
        } else {
            Logger.logWarn(
                    "Initializing WorkBookHandle failed: byte array does not contain a supported Excel WorkBook.");
            throw new InvalidFileException(
                    "byte array does not contian a supported Excel WorkBook.");
        }
    }

    protected void initWorkBookFactory() {
        myfactory = new WorkBookFactory();
    }

    /**
     * initialize from an XLSX/OOXML workbook.
     */
    private boolean initXLSX(String fname) {
        // do before parseNBind so can set myfactory & fname
        // set state vars for this workbookhandle
        this.initWorkBookFactory();

        myfactory.setDebugLevel(this.DEBUGLEVEL);
        myfactory.setFileName(this.name);

        if (plist != null)
            myfactory.register(plist); // register progress notifier
        try {
            // iterate sheets,inputting cell values, named ranges and
            // formula strings
            OOXMLReader oe = new OOXMLReader();
            WorkBookHandle bk = new WorkBookHandle();
            bk.removeAllWorkSheets();
            myfactory.setDebugLevel(this.DEBUGLEVEL);
            bk.DEBUGLEVEL = this.DEBUGLEVEL;
            oe.parseNBind(bk, fname);
            this.sheethandles = bk.sheethandles;
            this.mybook = bk.mybook;
        } catch (Exception e) {
            throw new WorkBookException(
                    "WorkBookHandle OOXML Read failed: " + e.toString(),
                    WorkBookException.UNSPECIFIED_INIT_ERROR, e);
        }

        mybook.setIsExcel2007(true);
        return true;
    }

    /**
     * do all initialization with a filename
     *
     * @param fname
     */
    protected void initFromFile(File fx) {
        String fname = fx.getPath();
        String finch = "";

        // handle csv import
        FileReader fincheck;
        try {
            fincheck = new FileReader(fx);
            if (fx.length() > 100) {
                char[] cbuf = new char[100];
                fincheck.read(cbuf);
                finch = new String(cbuf);
            }
            fincheck.close();

        } catch (FileNotFoundException e) {
            Logger.logErr("WorkBookHandle: Cannot open file " + fname + ": "
                    + e);
        } catch (Exception e1) {
            Logger.logErr("Invalid XLSX/OOXML File.");
        }
        this.name = fname; // 20081231 KSC: set here
        if (finch.toUpperCase().startsWith("PK")) { // it's a zip file... give
            // XLSX parsing a shot
            if (this.file != null)
                OOXMLAdapter.refreshPassThroughFiles(this);
            if (initXLSX(fname))
                return;
        }
        try {
            myLEOFile = new LEOFile(fx, this.DEBUGLEVEL);
        } catch (InvalidFileException ifx) {
            if ((finch.indexOf(",") > -1) && (finch.indexOf(",") > -1)) {
                // init a blank workbook
                this.initDefault();

                // map CSV into workbook
                try {
                    WorkSheetHandle sheet = getWorkSheet(0);
                    sheet.readCSV(new BufferedReader(new FileReader(fx)));
                    return;
                } catch (Exception e) {
                    throw new WorkBookException(
                            "Error encountered importing CSV: " + e.toString(),
                            WorkBookException.ILLEGAL_INIT_ERROR);
                }
            } else {
                throw ifx;
            }

        }
        if (myLEOFile.hasWorkBook()) {
            this.initFromLeoFile(myLEOFile);
        } else {
            // total failure to load
            Logger.logErr("Initializing WorkBookHandle failed: " + fname
                    + " does not contain a supported Excel WorkBook.");
            throw new InvalidFileException(
                    fname + " does not contian a supported Excel WorkBook.");
        }
    }

    /**
     * init the new WorkBookHandle
     */
    protected synchronized void initBytes(BlockByteReader blockByteReader) {
        this.initWorkBookFactory();

        if (plist != null)
            myfactory.register(plist); // register progress notifier
        myfactory.setDebugLevel(this.DEBUGLEVEL);

        mybook = (com.valkyrlabs.formats.XLS.WorkBook) myfactory
                .getWorkBook(blockByteReader, myLEOFile);

        if (dump_input != null) {
            try {
                dump_input.flush();
                dump_input.close();
                dump_input = null;
            } catch (Exception e) {
            }
        }
        this.postLoad();
    }

    /**
     * Handles tasks that need to occur after workbook has been loaded
     */
    void postLoad() {
        initHlinks();
        initMerges();
        mybook.initializeNames(); // must initialize name expressions AFTER
        // loading sheet records
        mybook.mergeMSODrawingRecords();
        mybook.initializeIndirectFormulas();
        initPivotCache(); // if any
    }

    void initMerges() {
        AbstractList<?> mergelookup = mybook.getMergecelllookup();
        for (int t = 0; t < mergelookup.size(); t++) {
            Mergedcells mc = (Mergedcells) mergelookup.get(t);
            mc.initCells(this);
        }
    }

    void initHlinks() {
        AbstractList<?> hlinklookup = mybook.getHlinklookup();
        for (int t = 0; t < hlinklookup.size(); t++) {
            Hlink hl = (Hlink) hlinklookup.get(t);
            hl.initCells(this);
        }
    }

    /**
     * reads in the pivot cache storage and parses the pivot cache records <br>
     * pivot cache(s) are used by pivot tables as data source storage
     */
    void initPivotCache() {
        if (myLEOFile.hasPivotCache()) {
            PivotCache pc = new PivotCache(); // grab any pivot caches
            try {
                pc.init(myLEOFile.getDirectoryArray(), this);
                mybook.setPivotCache(pc);
            } catch (Exception e) {

            }
        }
    }

    /**
     * Closes the WorkBook and releases resources.
     */
    @Override
    public void close() {
        try {
            if (myLEOFile != null)
                myLEOFile.shutdown();
            myLEOFile = null;
        } catch (Exception e) {
            if (DEBUGLEVEL > 3)
                Logger.logWarn("Closing Document: " + toString() + " failed: "
                        + e.toString());
        }
        if (mybook != null)
            mybook.close(); // clear out object refs to release memory
        mybook = null;
        myfactory = null;
        name = null;
        sheethandles = null;
        // Runtime.getRuntime().gc();
    }

    @Override
    protected void finalize() throws Throwable {
        close();
    }

    @Override
    public void reset() {
        initFromFile(new File(myLEOFile.getFileName()));
    }

    /**
     * Returns an array of handles to all of the WorkSheets in the Workbook.
     *
     * @return WorkSheetHandle[] Array of all WorkSheets in WorkBook
     */
    @Override
    public WorkSheetHandle[] getWorkSheets() {
        try {
            if (myfactory != null) {
                int numsheets = mybook.getNumWorkSheets();
                if (numsheets == 0)
                    throw new SheetNotFoundException(
                            "WorkBook has No Sheets.");
                WorkSheetHandle[] sheets = new WorkSheetHandle[numsheets];
                for (int i = 0; i < numsheets; i++) {
                    Boundsheet bs = mybook.getWorkSheetByNumber(i);
                    bs.setWorkBook(this.mybook);
                    sheets[i] = new WorkSheetHandle(bs, this);
                }
                return sheets;
            }
            return null;
        } catch (SheetNotFoundException a) {
            Logger.logWarn("getWorkSheets() failed: " + a);
            return null;
        }
    }

    /**
     * returns the handle to a WorkSheet by number.
     * <p>
     * Sheet 0 is the first Sheet.
     *
     * @param index of worksheet (ie: 0)
     * @return WorkSheetHandle the WorkSheet
     * @throws SheetNotFoundExceptionthe specified WorkSheet is not found in
     *                                   the WorkBook.
     */
    @Override
    public WorkSheetHandle getWorkSheet(int sheetnum) throws SheetNotFoundException {
        Boundsheet st = mybook.getWorkSheetByNumber(sheetnum);
        if (sheethandles.get(st.getSheetName()) != null)
            return sheethandles.get(st.getSheetName());
        else {
            WorkSheetHandle shth = new WorkSheetHandle(st, this);
            sheethandles.put(st.getSheetName(), shth);
            return shth;
        }
    }

    /**
     * returns the handle to a WorkSheet by name.
     *
     * @param String name of worksheet (ie: "Sheet1")
     * @return WorkSheetHandle the WorkSheet
     * @throws SheetNotFoundExceptionthe specified WorkSheet is not found in
     *                                   the WorkBook.
     */
    @Override
    public WorkSheetHandle getWorkSheet(String sheetName) throws SheetNotFoundException {
        if (sheethandles.get(sheetName) != null) {
            if (mybook.getWorkSheetByName(sheetName) != null) {
                return sheethandles.get(sheetName);
            } else {
                throw new SheetNotFoundException("Sheet " + sheetName + " not found");
            }
        }
        if (myfactory != null) {
            Boundsheet bs = mybook.getWorkSheetByName(sheetName);
            if (bs != null) {
                bs.setWorkBook(this.mybook);
                WorkSheetHandle ret = new WorkSheetHandle(bs, this);
                sheethandles.put(sheetName, ret);
                return ret;
            } else {
                throw new SheetNotFoundException(sheetName);
            }
        }
        throw new SheetNotFoundException("Cannot find WorkSheet " + sheetName);
    }

    /**
     * returns the active or selected worksheet tab
     *
     * @return WorkSheetHandle
     * @throws SheetNotFoundException{
     */
    public WorkSheetHandle getActiveSheet() throws SheetNotFoundException {
        return this.getWorkSheet(this.getWorkBook().getSelectedSheetNum());
    }

    /**
     * Returns a low-level WorkBook.
     * <p>
     * NOTE: The WorkBook class is NOT a part of the published API. Any of the
     * methods and/or variables on a WorkBook object are subject to change without
     * notice in new versions of OpenXLS.
     */
    @Override
    public com.valkyrlabs.formats.XLS.WorkBook getWorkBook() {
        return this.mybook;
    }

    /**
     * Set Encoding mode of new Strings added to file.
     * <p>
     * OpenXLS has 3 modes for handling the internal encoding of String data that is
     * added to the file.
     * <p>
     * OpenXLS can save space in the file if it knows that all characters in your
     * String data can be represented with a single byte (Compressed.)
     * <p>
     * If your String contains characters which need 2 bytes to represent (such as
     * Eastern-language characters) then it needs to be stored in an uncompressed
     * Unicode format.
     * <p>
     * OpenXLS can either automatically detect the mode for each String, or you can
     * set it explicitly. The auto mode is the most flexible but requires processing
     * overhead.
     * <p>
     * Default mode is WorkBookHandle.STRING_ENCODING_AUTO.
     * <p>
     * Valid Modes Are:
     * <p>
     * WorkBookHandle.STRING_ENCODING_AUTO Use if you are adding mixed Unicode and
     * non-unicode Strings and can accept the performance hit -slowest String adds
     * -optimal file size for mixed Strings WorkBookHandle.STRING_ENCODING_UNICODE
     * Use if all of your new Strings are Unicode - faster than AUTO -faster than
     * AUTO -largest file size WorkBookHandle.STRING_ENCODING_COMPRESSED Use if all
     * of your new Strings are non-Unicode and can have high-bytes compressed
     * -faster than AUTO -smallest file size
     *
     * @param int String Encoding Mode
     */
    @Override
    public void setStringEncodingMode(int mode) {
        mybook.setStringEncodingMode(mode);
    }

    /**
     * Set Duplicate String Handling Mode.
     *
     * <pre>
     * The Duplicate String Mode determines the behavior of
     * the String table when inserting new Strings.
     *
     * The String table shares a single entry for multiple
     * Cells containing the same string.  When multiple Cells
     * have the same value, they share the same underlying string.
     *
     * Changing the value of any one of the Cells will change
     * the value for any Cells sharing that reference.
     *
     * For this reason, you need to determine
     * the handling of new strings added to the sheet that
     * are duplicates of strings already in the table.
     *
     * If you will be changing the values of these
     * new Cells, you will need to set the Duplicate
     * String Mode to ALLOWDUPES.  If the string table
     * encounters a duplicate entry being added, it
     * will insert a duplicate that can then be subsequently
     * changed without affecting the other duplicate Cells.
     *
     * Valid Modes Are:
     *
     * WorkBookHandle.ALLOWDUPES - faster String inserts, larger file sizes,  changing Cells has no effect on dupe Cells
     *
     * WorkBookHandle.SHAREDUPES - slower inserts, dupe smaller file sizes, Cells share changes
     * </pre>
     *
     * @param int Duplicate String Handling Mode
     */
    @Override
    public void setDupeStringMode(int mode) {
        mybook.setDupeStringMode(mode);
    }

    /**
     * Copies an existing Chart to another WorkSheet
     *
     * @param chartname
     * @param sheetname
     */
    @Override
    public void copyChartToSheet(String chartname, String sheetname)
            throws ChartNotFoundException, SheetNotFoundException {
        mybook.copyChartToSheet(chartname, sheetname);
    }

    /**
     * Copies an existing Chart to another WorkSheet
     *
     * @param chart
     * @param sheet
     */
    @Override
    public void copyChartToSheet(ChartHandle chart, WorkSheetHandle sheet)
            throws ChartNotFoundException, SheetNotFoundException {
        mybook.copyChartToSheet(chart.getTitle(), sheet.getSheetName());
    }

    /**
     * Copy (duplicate) a worksheet in the workbook and add it to the end of the
     * workbook with a new name
     *
     * @param String the Name of the source worksheet;
     * @param String the Name of the new (destination) worksheet;
     * @return the new WorkSheetHandle
     */
    @Override
    public WorkSheetHandle copyWorkSheet(String SourceSheetName, String NewSheetName)
            throws SheetNotFoundException {
        try {
            mybook.copyWorkSheet(SourceSheetName, NewSheetName);
        } catch (Exception e) {
            throw new WorkBookException("Failed to copy WorkSheet: "
                    + SourceSheetName + ": " + e.toString(),
                    WorkBookException.RUNTIME_ERROR);
        }
        mybook.getRefTracker().clearPtgLocationCaches(NewSheetName);
        // update the merged cells (requires a WBH, that's why it's
        // here)
        WorkSheetHandle wsh = this.getWorkSheet(NewSheetName);
        if (wsh != null) {
            List<?> mc = wsh.getMysheet().getMergedCellsRecs();
            for (int i = 0; i < mc.size(); i++) {
                Mergedcells mrg = (Mergedcells) mc.get(i);
                if (mrg != null)
                    mrg.initCells(this);
            }
            // now conditional formats
            /*
             * mc = wsh.getMysheet().getConditionalFormats(); for (int
             * i=0;i<mc.size();i++)
             * { Condfmt mrg = (Condfmt)mc.get(i); if (mrg !=
             * null)mrg.initCells(this); }
             */
        }
        return wsh;
    }

    /**
     * Forces immediate recalculation of every formula in the workbook.
     *
     * @throws FunctionNotSupportedException if an unsupported function is used by
     *                                       any formula in the workbook
     * @see #forceRecalc()
     * @see #recalc()
     */
    @Override
    public void calculateFormulas() {
        markFormulasDirty();
        recalc();
    }

    /**
     * Marks every formula in the workbook as needing a recalc. This method does not
     * actually calculate formulas, for that use {@link #recalc()}.
     */
    public void markFormulasDirty() {
        Formula[] formulas = mybook.getFormulas();
        for (int idx = 0; idx < formulas.length; idx++)
            formulas[idx].clearCachedValue();
    }

    /**
     * Recalculates all dirty formulas in the workbook immediately.
     * <p>
     * You generally need not call this method. Dirty formulas will automatically be
     * recalculated when their values are queried. This method is only useful for
     * forcing calculation to occur at a certain time. In the case of functions such
     * as NOW() whose value is volatile the formula will still be recalculated every
     * time it is queried.
     *
     * @throws FunctionNotSupportedException if an unsupported function is used by
     *                                       any formula in the workbook
     * @see #markFormulasDirty()
     */
    public void recalc() {
        int calcmode = mybook.getCalcMode();
        mybook.setCalcMode(CALCULATE_AUTO); // ensure referenced functions are
        // calcualted as necesary!
        Formula[] formulas = mybook.getFormulas();
        for (int idx = 0; idx < formulas.length; idx++) {
            try {
                formulas[idx].clearCachedValue();
                formulas[idx].calculate();
            } catch (FunctionNotSupportedException fe) {
                Logger.logErr("WorkBookHandle.recalc:  Error calculating Formula "
                        + fe.toString());
            }
        }
        // KSC: Clear out lookup caches!
        this.getWorkBook().getRefTracker().clearLookupCaches();
        mybook.setCalcMode(calcmode); // reset
    }

    /**
     * Removes all of the WorkSheets from this WorkBook.
     * <p>
     * Bytes streamed from this WorkBook will create invalid Spreadsheet files
     * unless a WorkSheet(s) are added to it.
     * <p>
     * NOTE: A WorkBook with no sheets is *invalid* and will not open in Excel. You
     * must add sheets to this WorkBook for it to be valid.
     */
    @Override
    public void removeAllWorkSheets() {

        try {
            Object ob = this.mybook.getTabID().getTabIDs().get(0);
            this.mybook.getTabID().getTabIDs().clear();
            this.mybook.getTabID().getTabIDs().add(ob);
            this.mybook.getTabID().updateRecord();
        } catch (Exception ex) {
        }
        WorkSheetHandle[] ws = this.getWorkSheets();
        try {
            for (int x = 0; x < ws.length; x++) {
                try {
                    ws[x].remove();
                } catch (Exception e) {
                } // ignore the invalid WorkBook problem
            }
        } catch (Exception e) {
            // in case sheets already gone...
        }
        this.sheethandles.clear();
        this.mybook.closeSheets(); // replaced below with this

        /*
         * WHY ARE WE DOING THIS??? // init new book // save records
         * then reset to avoid
         * ByteStreamer.stream records expansion Object[] recs=
         * this.getWorkBook().getStreamer().getBiffRecords();
         *
         * // keep Excel 2007 status boolean isExcel2007=
         * this.getIsExcel2007();
         * WorkBookHandle ret = new WorkBookHandle(this.getBytes());
         * ret.setIsExcel2007(isExcel2007);
         *
         * this.getWorkBook().getStreamer().setBiffRecords(Arrays.
         * asList(recs));
         * this.mybook = ret.getWorkBook(); /
         **/

    }

    /**
     * Returns a WorkBookHandle containing an empty version of this WorkBook.
     * <p>
     * Use in conjunction with addSheetFromWorkBook() to create new output WorkBooks
     * containing various sheets from a master template.
     * <p>
     * ie: WorkBookHandle emptytemplate = this.getNoSheetWorkBook();
     * emptytemplate.addSheetFromWorkBook(this, "Sheet1", "TargetSheet");
     *
     * @return WorkBookHandle - the empty WorkBookHandle duplicate
     * @see addSheetFromWorkBook
     */
    @Override
    public WorkBookHandle getNoSheetWorkBook() {
        // to avoid ByteStreamer.stream records expansion
        Object[] recs = this.getWorkBook().getStreamer().getBiffRecords();
        byte[] gb = this.getBytes();
        WorkBookHandle ret = new WorkBookHandle(gb);
        this.getWorkBook().getStreamer().setBiffRecords(Arrays.asList(recs));
        ret.removeAllWorkSheets();
        return ret;
    }

    /**
     * Inserts a worksheet from a Source WorkBook.
     *
     * @param sourceBook      - the WorkBook containing the sheet to copy
     * @param sourceSheetName - the name of the sheet to copy
     * @param destSheetName   - the name of the new sheet in this workbook
     * @throws SheetNotFoundException{
     * @deprecated - use addWorkSheet(WorkSheetHandle sht, String NewSheetName){
     */
    @Override
    @Deprecated
    public boolean addSheetFromWorkBook(WorkBookHandle sourceBook, String sourceSheetName, String destSheetName)
            throws SheetNotFoundException {
        return this.addWorkSheet(sourceBook
                .getWorkSheet(sourceSheetName), destSheetName) != null;
    }

    /**
     * Inserts a worksheet from a Source WorkBook. Brings all string data and
     * formatting information from the source workbook.
     * <p>
     * Be aware this is programmatically creating a large amount of new formatting
     * information in the destination workbook. A higher performance option will
     * usually be using getNoSheetWorkbook and addSheetFromWorkBook.
     *
     * @param sourceBook      - the WorkBook containing the sheet to copy
     * @param sourceSheetName - the name of the sheet to copy
     * @param destSheetName   - the name of the new sheet in this workbook
     * @throws SheetNotFoundException{
     * @deprecated - use addWorkSheet(WorkSheetHandle sht, String NewSheetName){
     */
    @Deprecated
    public boolean addSheetFromWorkBookWithFormatting(WorkBookHandle sourceBook, String sourceSheetName,
            String destSheetName) throws SheetNotFoundException {
        return this.addWorkSheet(sourceBook
                .getWorkSheet(sourceSheetName), destSheetName) != null;
    }

    /**
     * Inserts a WorkSheetHandle from a separate WorkBookhandle into the current
     * WorkBookHandle.
     * <p>
     * copies charts, images, formats from source workbook
     * <p>
     * Worksheet will be the same name as in the source workbook. To add a custom
     * named worksheet use the addWorkSheet(WorkSheetHandle, String sheetname)
     * method
     *
     * @param WorkSheetHandle the source WorkSheetHandle;
     */
    public WorkSheetHandle addWorkSheet(WorkSheetHandle sourceSheet) {
        return this.addWorkSheet(sourceSheet, sourceSheet.getSheetName());
    }

    /**
     * Inserts a WorkSheetHandle from a separate WorkBookhandle into the current
     * WorkBookHandle.
     * <p>
     * copies charts, images, formats from source workbook
     *
     * @param WorkSheetHandle the source WorkSheetHandle;
     * @param String          the Name of the new (destination) worksheet;
     */
    @Override
    public WorkSheetHandle addWorkSheet(WorkSheetHandle sourceSheet, String NewSheetName) {
        sourceSheet.getSheet().populateForTransfer(); // copy all formatting +
        // images for this sheet
        List<?> chts = sourceSheet.getSheet().getCharts();
        for (int i = 0; i < chts.size(); i++) {
            Chart cxi = (Chart) chts.get(i);
            cxi.populateForTransfer();
        }
        byte[] bao = sourceSheet.getSerialBytes();
        try {
            mybook.addBoundsheet(bao, sourceSheet
                    .getSheetName(), NewSheetName,
                    StringTool
                            .stripPath(sourceSheet.getWorkBook()
                                    .getName()),
                    true);
            WorkSheetHandle wsh = this.getWorkSheet(NewSheetName);
            if (wsh != null) {
                List<?> mc = wsh.getMysheet().getMergedCellsRecs();
                for (int i = 0; i < mc.size(); i++) {
                    Mergedcells mrg = (Mergedcells) mc.get(i);
                    if (mrg != null)
                        mrg.initCells(this);
                }
            }

            return wsh;
        } catch (Exception e) {
            throw new WorkBookException(
                    "Failed to copy WorkSheet: " + e.toString(),
                    WorkBookException.RUNTIME_ERROR);
        }
    }

    /**
     * Utility method to copy a format handle from a separate WorkBookHandle to this
     * WorkBookHandle
     *
     * @param externalFormat - FormatHandle from an external WorkBookHandle
     * @return
     */
    public FormatHandle transferExternalFormatHandle(FormatHandle externalFormat) {
        Xf xf = externalFormat.getXf();
        FormatHandle newHandle = new FormatHandle(this);
        newHandle.addXf(xf);
        return newHandle;
    }

    /**
     * Creates a new Chart and places it at the end of the workbook
     *
     * @param String the Name of the newly created Chart
     * @return the new ChartHandle
     */
    public ChartHandle createChart(String name, WorkSheetHandle wsh) {
        if (wsh == null) {
            // this is a sheetless chart - TODO:
        }
        /*
         * a chart needs a supbook, externsheet, & MSO object in the
         * book stream. I
         * think this is due to the fact that the referenced series
         * are usually stored
         * in the fashon 'Sheet1!A4:B6' The sheet1 reference
         * requires a supbook, though
         * the reference is internal.
         */

        try {
            ObjectInputStream ois = new ObjectInputStream(
                    new ByteArrayInputStream(getPrototypeChart()));
            Chart newchart = (Chart) ois.readObject();
            newchart.setWorkBook(this.getWorkBook());
            if (this.getIsExcel2007())
                newchart = new OOXMLChart(newchart, this);
            mybook.addPreChart();
            mybook.addChart(newchart, name, wsh.getSheet());
            /*
             * add font recs if nec: for the default chart: default
             * chart text fonts are # 5
             * & 6 title # 7 axis # 8
             */
            ChartHandle bs = new ChartHandle(newchart, this);
            int nfonts = mybook.getNumFonts();
            while (nfonts < 8) { // ensure
                Font f = new Font("Arial", Font.PLAIN, 200);
                mybook.insertFont(f);
                nfonts++;
            }
            Font f = mybook.getFont(8); // axis title font
            if (f.toString()
                    .equals("Arial,400,200 java.awt.Color[r=0,g=0,b=0] font style:[falsefalsefalsefalse00]")) {
                // it's default text font -- change to default axis title
                // font
                f = new Font("Arial", Font.BOLD, 240);
                bs.setAxisFont(f);
            }
            f = mybook.getFont(7); // chart title font
            if (f.toString()
                    .equals("Arial,400,200 java.awt.Color[r=0,g=0,b=0] font style:[falsefalsefalsefalse00]")) {
                // it's default text font -- change to default title font
                f = new Font("Arial", Font.BOLD, 360);
                bs.setTitleFont(f);
            }
            bs.removeSeries(0); // remove the "dummied" series
            bs.setAxisTitle(ChartHandle.XAXIS, null); // remove default axis
            // titles, if any
            bs.setAxisTitle(ChartHandle.YAXIS, null); // ""
            return bs;
        } catch (Exception e) {
            Logger.logErr("Creating New Chart: " + name + " failed: " + e);
            return null;
        }
    }

    /**
     * delete an existing chart of the workbook
     *
     * @param chartname
     */
    public void deleteChart(String chartname, WorkSheetHandle wsh) throws ChartNotFoundException {
        try {
            mybook.deleteChart(chartname, wsh.getSheet());
        } catch (ChartNotFoundException e) {
            throw new ChartNotFoundException(
                    "Removing Chart: " + chartname + " failed: " + e);
        } catch (Exception e) {
            Logger.logErr("Removing Chart: " + chartname + " failed: " + e);
        }
    }

    /**
     * Returns the number of Sheets in this WorkBook
     *
     * @return int number of Sheets
     */
    public int getNumWorkSheets() {
        return mybook.getNumWorkSheets();
    }

    /**
     * Creates a new worksheet and places it at the specified position. The new
     * sheet will be inserted before the sheet currently at the given index. If the
     * given index is higher than the last index currently in use, the sheet will be
     * added to the end of the workbook and will receive an index one higher than
     * that of the current final sheet. If the given index is negative it will be
     * interpreted as 0.
     *
     * @param name     the name of the newly created worksheet
     * @param sheetpos the index at which the sheet should be inserted
     * @return the new WorkSheetHandle
     */
    public WorkSheetHandle createWorkSheet(String name, int sheetpos) {
        if (sheetpos > this.getNumWorkSheets())
            sheetpos = this.getNumWorkSheets();
        if (sheetpos < 0)
            sheetpos = 0;

        WorkSheetHandle s = this.createWorkSheet(name);
        s.setTabIndex(sheetpos);
        return s;
    }

    /**
     * Creates a new worksheet and places it at the end of the workbook.
     *
     * @param name the name of the newly created worksheet
     * @return the new WorkSheetHandle
     */
    @Override
    public WorkSheetHandle createWorkSheet(String name) {
        try {
            this.getWorkSheet(name);
            throw new WorkBookException(
                    "Attempting to add worksheet with duplicate name. " + name
                            + " already exists in " + this.toString(),
                    WorkBookException.RUNTIME_ERROR);
        } catch (SheetNotFoundException e) {
            // good!
        }

        Boundsheet bo = null;
        try {
            ObjectInputStream ois = new ObjectInputStream(
                    new ByteArrayInputStream(getPrototypeSheet()));
            bo = (Boundsheet) ois.readObject();
            mybook.addBoundsheet(bo, null, name, null, false);
            try {
                WorkSheetHandle bs = this.getWorkSheet(name);
                bs.setSelected(this.mybook.getNumWorkSheets() == 1); // it's the only sheet so select!
                return bs;
            } catch (SheetNotFoundException e) {
                Logger.logWarn("Creating New Sheet: " + name + " failed: " + e);
                return null;
            }
        } catch (Exception e) {
            Logger.logWarn("Error loading prototype sheet: " + e);
            return null;
        }
    }

    /**
     * Returns an array of all FormatHandles in the workbook
     */
    @Override
    public FormatHandle[] getFormats() {
        List<?> l = this.mybook.getXfrecs();
        FormatHandle[] formats = new FormatHandle[l.size()];
        Iterator<?> its = l.iterator();
        int i = 0;
        while (its.hasNext()) {
            Xf x = (Xf) its.next();
            // passing (this) with the format handle breaks the
            // relationship to the font.
            // if you need to pass it in we will have to handle it
            // differently
            try {
                formats[i] = new FormatHandle();
                formats[i].setWorkBook(this.getWorkBook());
                formats[i].setXf(x);
            } catch (Exception ex) {
            }
            i++;
        }
        return formats;
    }

    /**
     * Returns an array of all Conditional Formats in the workbook
     * <p>
     * these are formats referenced and used by the conditionally formatted ranges
     * in the workbook.
     *
     * @return
     */
    public FormatHandle[] getConditionalFormats() {
        // the idea is to create a fake IXFE for use by
        // sheetster to find formats
        // int cfxe = this.getWorkBook().getNumFormats() + 50000; //
        // there would have to
        // be 50k styles on the sheet to conflict here....
        // int cfxe = this.getWorkBook().getNumXfs() + 50000; //
        // there would have to be
        // 50k styles on the sheet to conflict here....

        List<FormatHandle> retl = new Vector<FormatHandle>();
        AbstractList<?> v = this.mybook.getSheetVect();

        Iterator<?> its = v.iterator();

        while (its.hasNext()) {
            Boundsheet shtx = (Boundsheet) its.next();
            List<?> fmtlist = shtx.getConditionalFormats();
            Iterator<?> ixa = fmtlist.iterator();
            while (ixa.hasNext()) {
                Condfmt cfm = (Condfmt) ixa.next();
                // cfm.initCells(this); // added!
                int cfxe = cfm.getCfxe();
                FormatHandle fz = new FormatHandle(cfm, this, cfxe, null);

                fz.setFormatId(cfxe);
                retl.add(fz);
                // cfm.setCfxe(cfxe);
                // cxfe++;
            }
        }
        FormatHandle[] formats = new FormatHandle[retl.size()];
        for (int t = 0; t < formats.length; t++)
            formats[t] = retl.get(t);
        return formats;
    }

    public String getWorkingDirectory() {
        // TODO Auto-generated method stub
        return null;
    }

    public void initSharedFormulas() {
        // TODO Auto-generated method stub

    }

    @Override
    public void writeBytes(OutputStream bbout) throws IOException {
        // TODO Auto-generated method stub
        throw new UnsupportedOperationException("Unimplemented method 'writeBytes'");
    }
}