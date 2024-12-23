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
package com.valkyrlabs.formats.OOXML;

import com.valkyrlabs.toolkit.Logger;
import org.xmlpull.v1.XmlPullParser;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;

/**
 * sheetView (Worksheet View)
 * <p>
 * A single sheet view definition. When more than 1 sheet view is defined in the file, it means that when opening
 * the workbook, each sheet view corresponds to a separate window within the spreadsheet application, where
 * each window is showing the particular sheet. containing the same workbookViewId value, the last sheetView
 * definition is loaded, and the others are discarded. When multiple windows are viewing the same sheet, multiple
 * sheetView elements (with corresponding workbookView entries) are saved.
 * <p>
 * parent:		sheetViews
 * children:  	pane, selection, pivotSelection
 */
// TODO: finish pivotSelection
public class SheetView implements OOXMLElement {

    private static final long serialVersionUID = 8750051341951797617L;
    private HashMap<String, Object> attrs = new HashMap<String, Object>();
    private Pane pane = null;
    private ArrayList<Selection> selections = new ArrayList<Selection>();

    public SheetView() {

    }

    public SheetView(HashMap<String, Object> attrs, Pane p, ArrayList<Selection> selections) {
        this.attrs = attrs;
        this.pane = p;
        this.selections = selections;
    }

    public SheetView(SheetView s) {
        this.attrs = s.attrs;
        this.pane = s.pane;
        this.selections = s.selections;
    }

    public static OOXMLElement parseOOXML(XmlPullParser xpp) {
        HashMap<String, Object> attrs = new HashMap<String, Object>();
        Pane p = null;
        ArrayList<Selection> selections = new ArrayList<Selection>();
        try {
            int eventType = xpp.getEventType();
            while (eventType != XmlPullParser.END_DOCUMENT) {
                if (eventType == XmlPullParser.START_TAG) {
                    String tnm = xpp.getName();
                    if (tnm.equals("sheetView")) {        // get attributes
                        for (int i = 0; i < xpp.getAttributeCount(); i++) {
                            attrs.put(xpp.getAttributeName(i), xpp.getAttributeValue(i));
                        }
                    } else if (tnm.equals("pane")) {
                        p = Pane.parseOOXML(xpp);
                    } else if (tnm.equals("selection")) {
                        selections.add(Selection.parseOOXML(xpp));
                    }
                } else if (eventType == XmlPullParser.END_TAG) {
                    String endTag = xpp.getName();
                    if (endTag.equals("sheetView")) {
                        break;
                    }
                }
                eventType = xpp.next();
            }
        } catch (Exception e) {
            Logger.logErr("sheetView.parseOOXML: " + e.toString());
        }
        SheetView s = new SheetView(attrs, p, selections);
        return s;
    }

    public String getOOXML() {
        StringBuffer ooxml = new StringBuffer();
        ooxml.append("<sheetView");
        // attributes
        Iterator<String> i = attrs.keySet().iterator();
        while (i.hasNext()) {
            String key = i.next();
            String val = (String) attrs.get(key);
            ooxml.append(" " + key + "=\"" + val + "\"");
        }
        ooxml.append(">");
        if (pane != null) ooxml.append(pane.getOOXML());
        if (selections.size() > 0) {
            for (int j = 0; j < selections.size(); j++) {
                ooxml.append(selections.get(j).getOOXML());
            }
        }
        ooxml.append("</sheetView>");
        return ooxml.toString();
    }

    public OOXMLElement cloneElement() {
        return new SheetView(this);
    }


    /**
     * return the attribute value for key null if not found
     *
     * @param key
     * @return
     */
    public Object getAttr(String key) {
        return this.attrs.get(key);
    }

    /**
     * set the atttribute value for key to val
     *
     * @param key
     * @param val
     */
    public void setAttr(String key, Object val) {
        this.attrs.put(key, val);
    }

    /**
     * remove a previously set attribute, if found
     *
     * @param key
     */
    public void removeAttr(String key) {
        this.attrs.remove(key);
    }

    public void removeSelection() {
        this.removeAttr("tabSelected");
        selections = new ArrayList<Selection>();
    }

    /**
     * return the attribute value for key
     * in String form "" if not found
     *
     * @param key
     * @return
     */
    public String getAttrS(String key) {
        Object o = this.attrs.get(key);
        if (o == null)
            return "";
        return o.toString();
    }

}

/**
 * pane (View Pane)
 * Worksheet view pane
 * <p>
 * parent:  sheetView, customSheetView
 * children: none
 */
class Pane implements OOXMLElement {
    private static final long serialVersionUID = 5570779997661362205L;
    private HashMap<String, String> attrs = null;

    public Pane(HashMap<String, String> attrs) {
        this.attrs = attrs;
    }

    public Pane(Pane p) {
        this.attrs = p.attrs;
    }


    public static Pane parseOOXML(XmlPullParser xpp) {
        HashMap<String, String> attrs = new HashMap<String, String>();
        try {
            int eventType = xpp.getEventType();
            while (eventType != XmlPullParser.END_DOCUMENT) {
                if (eventType == XmlPullParser.START_TAG) {
                    String tnm = xpp.getName();
                    if (tnm.equals("pane")) {        // get attributes
                        for (int i = 0; i < xpp.getAttributeCount(); i++) {
                            attrs.put(xpp.getAttributeName(i), xpp.getAttributeValue(i));
                        }
                    }
                } else if (eventType == XmlPullParser.END_TAG) {
                    String endTag = xpp.getName();
                    if (endTag.equals("pane")) {
                        break;
                    }
                }
                eventType = xpp.next();
            }
        } catch (Exception e) {
            Logger.logErr("pane.parseOOXML: " + e.toString());
        }
        Pane p = new Pane(attrs);
        return p;
    }

    public String getOOXML() {
        StringBuffer ooxml = new StringBuffer();
        ooxml.append("<pane");
        // attributes
        Iterator<String> i = attrs.keySet().iterator();
        while (i.hasNext()) {
            String key = i.next();
            String val = attrs.get(key);
            ooxml.append(" " + key + "=\"" + val + "\"");
        }
        ooxml.append("/>");
        return ooxml.toString();
    }

    public OOXMLElement cloneElement() {
        return new Pane(this);
    }
}

/**
 * selection (Selection)
 * Worksheet view selection.
 * <p>
 * parent: 	 sheetView, customSheetView
 * children: none
 */
class Selection implements OOXMLElement {

    private static final long serialVersionUID = -5411798327743116154L;
    private HashMap<String, String> attrs = null;

    public Selection(HashMap<String, String> attrs) {
        this.attrs = attrs;
    }

    public Selection(Selection s) {
        this.attrs = s.attrs;
    }

    public static Selection parseOOXML(XmlPullParser xpp) {
        HashMap<String, String> attrs = new HashMap<String, String>();
        try {
            int eventType = xpp.getEventType();
            while (eventType != XmlPullParser.END_DOCUMENT) {
                if (eventType == XmlPullParser.START_TAG) {
                    String tnm = xpp.getName();
                    if (tnm.equals("selection")) {        // get attributes
                        for (int i = 0; i < xpp.getAttributeCount(); i++) {
                            attrs.put(xpp.getAttributeName(i), xpp.getAttributeValue(i));
                        }
                    }
                } else if (eventType == XmlPullParser.END_TAG) {
                    String endTag = xpp.getName();
                    if (endTag.equals("selection")) {
                        break;
                    }
                }
                eventType = xpp.next();
            }
        } catch (Exception e) {
            Logger.logErr("selection.parseOOXML: " + e.toString());
        }
        Selection s = new Selection(attrs);
        return s;
    }

    public String getOOXML() {
        StringBuffer ooxml = new StringBuffer();
        ooxml.append("<selection");
        // attributes
        Iterator<String> i = attrs.keySet().iterator();
        while (i.hasNext()) {
            String key = i.next();
            String val = attrs.get(key);
            ooxml.append(" " + key + "=\"" + val + "\"");
        }
        ooxml.append("/>");
        return ooxml.toString();
    }

    public OOXMLElement cloneElement() {
        return new Selection(this);
    }
}

