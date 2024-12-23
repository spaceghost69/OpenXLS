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
package com.valkyrlabs.formats.cellformat;

import com.valkyrlabs.OpenXLS.Cell;
import com.valkyrlabs.OpenXLS.CellHandle;
import com.valkyrlabs.OpenXLS.DateConverter;

import java.text.SimpleDateFormat;
import java.util.Calendar;

public class DateCellFormat
        extends SimpleDateFormat
        implements CellFormat {
    private static final long serialVersionUID = 1896075041723437260L;

    private final String text_format;

    DateCellFormat(String date, String text) {
        super(date);
        this.text_format = text;
    }

    public String format(Cell cell) {
        // make sure to return the empty string for blank cells
        // getting the calendar coerces to double and thus gets zero
        if ((cell instanceof CellHandle && ((CellHandle) cell).isBlank())
                || "".equals(cell.getVal()))
            return "";

        if (cell.getCellType() == Cell.TYPE_STRING) {
            return String.format(this.text_format, cell.getVal());
        }

        Calendar date = DateConverter.getCalendarFromCell(cell);
        return this.format(date.getTime());
    }
}
