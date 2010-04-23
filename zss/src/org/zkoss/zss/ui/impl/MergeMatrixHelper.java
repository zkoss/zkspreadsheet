/* MergeMatrixHelper.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Mar 18, 2008 1:08:11 PM     2008, Created by Dennis.Chen
}}IS_NOTE

Copyright (C) 2007 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
	This program is distributed under GPL Version 2.0 in the hope that
	it will be useful, but WITHOUT ANY WARRANTY.
}}IS_RIGHT
*/
package org.zkoss.zss.ui.impl;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.LinkedList;
import java.util.List;
import java.util.Map;

//import org.zkoss.zss.model.Sheet;
import org.zkoss.zss.ui.Rect;

import org.apache.poi.ss.usermodel.Sheet;

/**
 * Because current implementation of spreadsheet only support horizontal merged cell,
 * So, use this helper to help speradsheet control merged cell in horizontal only
 * @author Dennis.Chen
 *
 */
public class MergeMatrixHelper {

	
	private Map _leftTopIndex = new HashMap(5);
	private Map _mergeByIndex = new HashMap(20);
	private List _mergeRanges = new LinkedList();
	
	private int _frozenRow;
	private int _frozenCol;//don't care this for now.
	private SequenceId _mergeId = new SequenceId(0,1);
	/**
	 * @param mergeRange List of merge range
	 * @param frozenRow
	 * @param frozenCol
	 */
	public MergeMatrixHelper(List mergeRange,int frozenRow,int frozenCol){
		Iterator iter = mergeRange.iterator();
		while(iter.hasNext()){
			int[] r = (int[])iter.next();
			int left = r[0];
			int top = r[1];
			int right = r[2];
			int bottom = r[3];
			//System.out.println("Merge:"+(count++)+" l:"+left+",t:"+top+",r:"+right+",b:"+bottom);
			for(int i=top;i<=bottom;i++){
				MergedRect block = new MergedRect(_mergeId.next(),left,i,right,i);
				_mergeRanges.add(block);
			}
		}
		
		this._frozenRow = frozenRow;
		this._frozenCol = frozenCol;
		
		rebuildIndex();
	}
	
	private void rebuildIndex(){
		_leftTopIndex = new HashMap(5);
		_mergeByIndex = new HashMap(20);
		
		Iterator iter = _mergeRanges.iterator();
		while(iter.hasNext()){
			MergedRect block= (MergedRect)iter.next();
			int left = block.getLeft();
			int top = block.getTop();
			int right = block.getRight();
			_leftTopIndex.put(top+"_"+left,block);
			for(int j=left;j<=right;j++){
				_mergeByIndex.put(top+"_"+j, block);
			}
		}
	}
	
	
	public void update(int frozenRow,int frozenCol){
		this._frozenRow = frozenRow;
		this._frozenCol = frozenCol;
	}
	
	/**
	 * Check is (row,col) in one of merge range's left-top 
	 */
	public boolean isMergeRangeLeftTop(int row,int col){
		return _leftTopIndex.get(row+"_"+col)==null?false:true;
	}
	
	/**
	 * Get a merged range which contains cell(row,col)
	 */
	public MergedRect getMergeRange(int row,int col){
		MergedRect range = (MergedRect)_mergeByIndex.get(row+"_"+col);
		return range;
	}
	
	/**
	 * Get all merged range. 
	 * @return a list which contains all merged range
	 */
	public List getRanges(){
		return _mergeRanges;
	}
	
	/**
	 * Get merged range which contains col
	 * @param col column index
	 * @return a list which contains merged range
	 */
	public List getRangesByColumn(int col){
		Iterator iter = _mergeRanges.iterator();
		List result = new ArrayList();
		while(iter.hasNext()){
			Rect rect = (Rect)iter.next();
			int left = rect.getLeft();
			int right = rect.getRight();
			if(left<=col && right>=col){
				result.add(rect);
			}
		}
		return result;
	}

	public int getRightConnectedColumn(int col, int top, int bottom) {
		int size = _mergeRanges.size();
		List result = new ArrayList();
		Rect rect;
		for(int i=0;i<size;i++){
			rect = (MergedRect)_mergeRanges.get(i);
			if(rect.getTop()>_frozenRow && (rect.getTop()<top || rect.getBottom()>bottom)){
				continue;
			}
			result.add(rect);
		}
		
		boolean conti = true;
		while(conti){
			conti = false;
			size = result.size();
			for(int i=0;i<size;i++){
				rect = (MergedRect)result.get(i);
				if(rect.getRight()>col && rect.getLeft()<=col){
					col = rect.getRight();
					conti = true;
					result.remove(i);
					break;
				}
			}
		}
		return col;
	}

	public int getLeftConnectedColumn(int col, int top, int bottom) {
		int size = _mergeRanges.size();
		List result = new ArrayList();
		Rect rect;
		for(int i=0;i<size;i++){
			rect = (MergedRect)_mergeRanges.get(i);
			if(rect.getTop()>_frozenRow && (rect.getTop()<top || rect.getBottom()>bottom)){
				continue;
			}
			result.add(rect);
		}
		
		boolean conti = true;
		while(conti){
			conti = false;
			size = result.size();
			for(int i=0;i<size;i++){
				rect = (MergedRect)result.get(i);
				if(rect.getLeft()<col && rect.getRight()>=col){
					col = rect.getLeft();
					conti = true;
					result.remove(i);
					break;
				}
			}
		}
		return col;
	}

	public void updateMergeRange(int oleft, int otop, int oright, int obottom, int left, int top, int right,
			int bottom, List toadd, List torem) {
		for(int i=otop;i<=obottom;i++){
			MergedRect mblock = getMergeRange(i,oleft);
			if(mblock!=null){
				torem.add(mblock);
				_mergeRanges.remove(mblock);
			}
		}
		for(int i=top;i<=bottom;i++){
			MergedRect mblock = new MergedRect(_mergeId.next(),left,i,right,i);
			toadd.add(mblock);
			_mergeRanges.add(mblock);
		}
		rebuildIndex();
	}

	public void deleteMergeRange(int left, int top, int right, int bottom, List torem) {
		for(int i=top;i<=bottom;i++){
			MergedRect mblock = getMergeRange(i,left);
			if(mblock!=null){
				torem.add(mblock);
				_mergeRanges.remove(mblock);
			}
		}
		rebuildIndex();
	}

	public void addMergeRange(int left, int top, int right, int bottom, List toadd,List torem) {
		for(int i=top;i<=bottom;i++){
			MergedRect mblock = getMergeRange(i,left);
			if(mblock!=null){
				torem.add(mblock);
				_mergeRanges.remove(mblock);
			}
		}
		for(int i=top;i<=bottom;i++){
			MergedRect mblock = new MergedRect(_mergeId.next(),left,i,right,i);
			toadd.add(mblock);
			_mergeRanges.add(mblock);
		}
		rebuildIndex();
	}

	public void deleteAffectedMergeRangeByColumn(int col,List removed) {
		for(Iterator iter = _mergeRanges.iterator();iter.hasNext();){
			MergedRect block = (MergedRect)iter.next();
			int right = block.getRight();
			if(right<col) continue;
			removed.add(block);
		}
		for(Iterator iter = removed.iterator();iter.hasNext();){
			MergedRect block = (MergedRect)iter.next();
			_mergeRanges.remove(block);
		}
		rebuildIndex();
	}

	public void deleteAffectedMergeRangeByRow(int row,List removed) {
		for(Iterator iter = _mergeRanges.iterator();iter.hasNext();){
			MergedRect block = (MergedRect)iter.next();
			int bottom = block.getBottom();
			if(bottom<row) continue;
			removed.add(block);
		}
		for(Iterator iter = removed.iterator();iter.hasNext();){
			MergedRect block = (MergedRect)iter.next();
			_mergeRanges.remove(block);
		}
		rebuildIndex();
	}
	
}
