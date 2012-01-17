/* MergeAggregation.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Jan 16, 2012 10:34:45 AM , Created by sam
}}IS_NOTE

Copyright (C) 2012 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
package org.zkoss.zss.ui.impl;

import java.util.HashMap;
import java.util.Iterator;
import java.util.LinkedHashMap;

import org.zkoss.json.JSONArray;
import org.zkoss.json.JSONObject;

/**
 * @author sam
 *
 */
public class MergeAggregation {

	private MergeMatrixHelper _helper;
	
	private StringAggregation _mergeCSS = new StringAggregation();
	private HashMap<String, Integer> _cssIndex = new HashMap<String, Integer>();
	
	/**
	 * Key: merge id of MergedRect
	 * Value: merge index of the aggregation
	 */
	private LinkedHashMap<Integer, Integer> _mergeIds = new LinkedHashMap<Integer, Integer>();
	
	public MergeAggregation (MergeMatrixHelper helper) {
		_helper = helper;
	}
	
	/**
	 * Add merge info of cell
	 * 
	 * @param int the row number
	 * @param int the column number
	 * @return index merge info index. Returns -1 if there's no merge info
	 */
	public MergeIndex add(int row, int col) {
		int index = -1; //index of merge id
		MergedRect block = _helper.getMergeRange(row, col);
		if (block != null) {
			Integer id = Integer.valueOf(block.getId());
			if (_mergeIds.containsKey(id)) {
				index = _mergeIds.get(id);
			} else {
				index = _mergeIds.size();
				_mergeIds.put(id, index);
			}
			
			String css = "";
			boolean isLeftTop = _helper.isMergeRangeLeftTop(row, col);
			if (isLeftTop) {
				css = "zsmerge" + id;
			} else if (block.getTop() == row) {
				css = "zsmergee";
			} else {
				css = "zsmergeeu";
			}
			return new MergeIndex(index, _mergeCSS.add(css));
		}
		return null;
	}
	
	/**
	 * 
	 * <ul>
	 * 	<li>r: merge rects</li>
	 *  <li>cs: merge CSS</li>
	 * </ul>
	 * 
	 * @return
	 */
	public JSONObject getJSONObject() {
		JSONObject result = new JSONObject();
		result.put("r", getMerge());
		result.put("cs", getMergeCSS());
		return result;
	}
	
	private JSONArray getMerge() {
		JSONArray ary = new JSONArray();
		
		Iterator<Integer> iter = _mergeIds.keySet().iterator();
		while (iter.hasNext()) {
			Integer id = iter.next(); //merge id
			ary.add(newMergeAttributes(_helper.getMergedRect(id)));
		}
		
		return ary;
	}
	
	/**
	 * Merge attributes
	 * 
	 * <ul>
	 * 	<li>i: id</li>
	 *  <li>l: left</li>
	 *  <li>t: top</li>
	 *  <li>r: right</li>
	 *  <li>b: bottom</li>
	 * </ul>
	 * 
	 * @param rect
	 * @return
	 */
	private HashMap<String, Integer> newMergeAttributes(MergedRect rect) {
		HashMap<String, Integer> attrs = new HashMap<String, Integer>(5); //id, left, right, top and bottom attributes
		attrs.put("i", rect.getId());
		attrs.put("l", rect.getLeft());
		attrs.put("t", rect.getTop());
		attrs.put("r", rect.getRight());
		attrs.put("b", rect.getBottom());
		return attrs;
	}
	
	private JSONArray getMergeCSS() {
		return _mergeCSS.getJSONArray();
	}
	
	public class MergeIndex {
		int mergeId;
		int mergeCSSId;
		
		MergeIndex (int id, int cssId) {
			mergeId = id;
			mergeCSSId = cssId;
		}
		
		public int getMergeId() {
			return mergeId;
		}
		
		public int getMergeCSSId() {
			return mergeCSSId;
		}
	}
}
