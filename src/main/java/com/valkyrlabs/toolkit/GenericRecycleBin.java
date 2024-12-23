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
package com.valkyrlabs.toolkit;

import java.util.Collection;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.Stack;
import java.util.Vector;

/**
 * A recycling cache, items are checked at intervals
 * 
 * 
 *
 */
public abstract class GenericRecycleBin extends java.lang.Thread
		implements Map<Object, Object>, com.valkyrlabs.toolkit.RecycleBin {
	protected Map<Object, Object> map = new java.util.HashMap();
	protected Vector<Object> active = new Vector<Object>();
	protected Stack<Recyclable> spares = new Stack<Recyclable>();

	/**
	 * add an item
	 */
	public void addItem(Recyclable r) throws RecycleBinFullException {
		if ((MAXITEMS == -1) || (map.size() < MAXITEMS)) {
			addItem(Integer.valueOf(map.size()), r);
		} else {
			throw new RecycleBinFullException();
		}
	}

	/**
	 * returns number of items in cache
	 * 
	 * 
	 * @return
	 */
	public int getNumItems() {
		return active.size();
	}

	public void addItem(Object key, Recyclable r) throws RecycleBinFullException {
		// recycle();
		if ((MAXITEMS == -1) || (map.size() < MAXITEMS)) {
			active.add(r);
			map.put(key, r);

		} else {
			throw new RecycleBinFullException();
		}
	}

	/**
	 * iterate all active items and try to recycle
	 */
	public synchronized void recycle() {
		Recyclable[] rs = new Recyclable[active.size()];
		active.copyInto(rs);
		for (int t = 0; t < rs.length; t++) {
			try {
				Recyclable rb = rs[t];
				if (!rb.inUse()) {
					// recycle it
					rb.recycle();

					// remove from active and lookup
					active.remove(rb);
					map.remove(rb);

					// put in spares
					spares.push(rb);

				}
			} catch (Exception ex) {
				Logger.logErr("recycle failed", ex);
			}
		}

	}

	public void empty() {
		map.clear();
		active.clear();
	}

	public synchronized List<Object> getAll() {
		return active;
	}

	/**
	 * returns a new or recycled item from the spares pool
	 * 
	 * 
	 * @see com.valkyrlabs.toolkit.RecycleBin#getItem()
	 */
	public synchronized Recyclable getItem() throws RecycleBinFullException {
		Recyclable active = null;
		// spares contains the recycled
		if (spares.size() > 0) {
			active = spares.pop();
			addItem(active);
			return active;
		}
		recycle();

		// technically infinite loop until exception thrown
		return getItem();
	}

	protected int MAXITEMS = -1; // no limit is default

	/**
	 * max number of items to be put in this bin.
	 * 
	 */
	public void setMaxItems(int i) {
		MAXITEMS = i;
	}

	public int getMaxItems() {
		return MAXITEMS;
	}

	public int getSpareCount() {
		return spares.size();
	}

	public GenericRecycleBin() {
	}

	public void clear() {
		map.clear();
		active.clear();
	}

	public boolean containsKey(Object key) {
		return map.containsKey(key);

	}

	public boolean containsValue(Object value) {
		return map.containsValue(value);
	}

	public Set entrySet() {
		return map.entrySet();
	}

	@Override
	public boolean equals(Object o) {
		return map.equals(o);
	}

	public Object get(Object key) {
		return map.get(key);
	}

	@Override
	public int hashCode() {
		return map.hashCode();
	}

	public boolean isEmpty() {
		return map.isEmpty();
	}

	public Set<Object> keySet() {
		return map.keySet();
	}

	public Object put(Object arg0, Object arg1) {
		active.add(arg1);
		return map.put(arg0, arg1);
	}

	public void putAll(Map<?, ?> arg0) {
		active.addAll(arg0.entrySet());
		map.putAll(arg0);
	}

	public Object remove(Object key) {
		active.remove(map.get(key));
		return map.remove(key);
	}

	public int size() {
		return map.size();
	}

	@Override
	public String toString() {
		return map.toString();
	}

	public Collection<Object> values() {
		return map.values();
	}

	public java.util.Map getMap() {
		return map;
	}

	public void setMap(java.util.HashMap _map) {
		map = _map;
	}

}