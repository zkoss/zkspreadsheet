/* CellFetchCommandHelper.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		January 10, 2008 03:10:40 PM , Created by Dennis.Chen
}}IS_NOTE

Copyright (C) 2007 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
	This program is distributed under Lesser GPL Version 2.1 in the hope that
	it will be useful, but WITHOUT ANY WARRANTY.
}}IS_RIGHT
*/
package org.zkoss.zss.ui.au.in;


import java.util.ArrayList;
import java.util.List;
import java.util.Map;

import org.zkoss.lang.Objects;
import org.zkoss.poi.ss.usermodel.Cell;
import org.zkoss.poi.ss.usermodel.CellStyle;
import org.zkoss.poi.ss.usermodel.Hyperlink;
import org.zkoss.poi.ss.usermodel.RichTextString;
import org.zkoss.util.logging.Log;
import org.zkoss.zk.au.AuRequest;
import org.zkoss.zk.mesg.MZk;
import org.zkoss.zk.ui.Component;
import org.zkoss.zk.ui.UiException;
import org.zkoss.zss.model.Worksheet;
import org.zkoss.zss.model.FormatText;
import org.zkoss.zss.model.impl.BookHelper;
import org.zkoss.zss.model.impl.SheetCtrl;
import org.zkoss.zss.ui.Spreadsheet;
import org.zkoss.zss.ui.impl.CellFormatHelper;
import org.zkoss.zss.ui.impl.HeaderPositionHelper;
import org.zkoss.zss.ui.impl.JSONObj;
import org.zkoss.zss.ui.impl.MergeMatrixHelper;
import org.zkoss.zss.ui.impl.MergedRect;
import org.zkoss.zss.ui.impl.Utils;
import org.zkoss.zss.ui.impl.HeaderPositionHelper.HeaderPositionInfo;
import org.zkoss.zss.ui.sys.SpreadsheetCtrl;
import org.zkoss.zss.ui.sys.SpreadsheetInCtrl;


/**
 * A Command Helper for (client to server) for fetch data back
 * @author Dennis.Chen
 *
 */
public class CellFetchCommandHelper{
	private static final Log log = Log.lookup(CellFetchCommandHelper.class);
	
	private Spreadsheet _spreadsheet;
	private SpreadsheetCtrl _ctrl;
	HeaderPositionHelper _rowHelper;
	HeaderPositionHelper _colHelper;
	private MergeMatrixHelper _mergeMatrix;
	private boolean _hidecolhead;
	private boolean _hiderowhead;
	private int _lastleft;
	private int _lastright;
	private int _lasttop;
	private int _lastbottom;
	
	private void responseDataBlock(String postfix, String token, String sheetid, String result) {
		//bug 1953830 Unnecessary command was sent and break the processing
		//use smartUpdate to instead
		//_spreadsheet.response(null, new org.zkoss.zss.ui.au.out.AuDataBlock(_spreadsheet,token,sheetid,result));
		
		//to avoid response be override in smartUpdate, I use a count-postfix
		//_spreadsheet.smartUpdateValues("dblock_"+Utils.nextUpdateId(),new Object[]{token,sheetid,result});
		
		_spreadsheet.smartUpdate(postfix != null ? "dataBlockUpdate" + postfix : "dataBlockUpdate", new String[] {token, sheetid, result});
	}
	
	//-- super --//
	protected void process(AuRequest request) {
		final Component comp = request.getComponent();
		if (comp == null)
			throw new UiException(MZk.ILLEGAL_REQUEST_COMPONENT_REQUIRED, this);
		final Map data = request.getData();
		if (data == null || data.size() != 20)
			throw new UiException(MZk.ILLEGAL_REQUEST_WRONG_DATA,
				new Object[] {Objects.toString(data), this});
		
		_spreadsheet = ((Spreadsheet)comp);
		if(_spreadsheet.isInvalidated()) return;//since it is invalidate, i don't need to update
		final Worksheet selSheet = _spreadsheet.getSelectedSheet();
		final String sheetId = (String) data.get("sheetId");
		if (selSheet == null || !sheetId.equals(((SheetCtrl)selSheet).getUuid())) { //not current selected sheet, skip.
			return;
		}
		
		_ctrl = ((SpreadsheetCtrl)_spreadsheet.getExtraCtrl());
		_hidecolhead = _spreadsheet.isHidecolumnhead();
		_hiderowhead = _spreadsheet.isHiderowhead();
		String token = (String) data.get("token");
		
		_rowHelper = _ctrl.getRowPositionHelper(sheetId);
		_colHelper = _ctrl.getColumnPositionHelper(sheetId);
		
		Worksheet sheet = _spreadsheet.getSelectedSheet();
		if(!Utils.getSheetUuid(sheet).equals(sheetId)) return;
		
		_mergeMatrix = _ctrl.getMergeMatrixHelper(sheet);
		
		String type = (String) data.get("type"); 
		String direction = (String) data.get("direction");
		
		
		int dpWidth = (Integer)data.get("dpWidth");//pixel value of data panel width
		int dpHeight = (Integer)data.get("dpHeight");//pixel value of data panel height
		int viewWidth = (Integer)data.get("viewWidth");//pixel value of view width(scrollpanel.clientWidth)
		int viewHeight = (Integer)data.get("viewHeight");//pixel value of value height
		
		int blockLeft = (Integer)data.get("blockLeft");
		int blockTop = (Integer)data.get("blockTop"); 
		int blockRight = (Integer)data.get("blockRight");// + blockLeft - 1;
		int blockBottom = (Integer)data.get("blockBottom");// + blockTop - 1;;
		
		int fetchLeft = (Integer)data.get("fetchLeft");
		int fetchTop = (Integer)data.get("fetchTop"); 
		int fetchWidth = (Integer)data.get("fetchWidth");
		int fetchHeight = (Integer)data.get("fetchHeight");
		
		int rangeLeft = (Integer)data.get("rangeLeft");//visible range
		int rangeTop = (Integer)data.get("rangeTop"); 
		int rangeRight = (Integer)data.get("rangeRight");
		int rangeBottom = (Integer)data.get("rangeBottom");
		
		try{
			if("jump".equals(type)){
				String result = null;
				if("east".equals(direction)){
					result = jump("E",(Spreadsheet)comp,sheetId,sheet,type,dpWidth,dpHeight,viewWidth,viewHeight,blockLeft,blockTop,blockRight,blockBottom,fetchLeft,fetchTop,rangeLeft,rangeTop,rangeRight,rangeBottom);
				}else if("south".equals(direction)){
					result = jump("S",(Spreadsheet)comp,sheetId,sheet,type,dpWidth,dpHeight,viewWidth,viewHeight,blockLeft,blockTop,blockRight,blockBottom,fetchLeft,fetchTop,rangeLeft,rangeTop,rangeRight,rangeBottom);
				}else if("west".equals(direction)){
					result = jump("W",(Spreadsheet)comp,sheetId,sheet,type,dpWidth,dpHeight,viewWidth,viewHeight,blockLeft,blockTop,blockRight,blockBottom,fetchLeft,fetchTop,rangeLeft,rangeTop,rangeRight,rangeBottom);
				}else if("north".equals(direction)){
					result = jump("N",(Spreadsheet)comp,sheetId,sheet,type,dpWidth,dpHeight,viewWidth,viewHeight,blockLeft,blockTop,blockRight,blockBottom,fetchLeft,fetchTop,rangeLeft,rangeTop,rangeRight,rangeBottom);
				}else if("westnorth".equals(direction)){
					result = jump("WN",(Spreadsheet)comp,sheetId,sheet,type,dpWidth,dpHeight,viewWidth,viewHeight,blockLeft,blockTop,blockRight,blockBottom,fetchLeft,fetchTop,rangeLeft,rangeTop,rangeRight,rangeBottom);
				}else if("eastnorth".equals(direction)){
					result = jump("EN",(Spreadsheet)comp,sheetId,sheet,type,dpWidth,dpHeight,viewWidth,viewHeight,blockLeft,blockTop,blockRight,blockBottom,fetchLeft,fetchTop,rangeLeft,rangeTop,rangeRight,rangeBottom);
				}else if("westsouth".equals(direction)){
					result = jump("WS",(Spreadsheet)comp,sheetId,sheet,type,dpWidth,dpHeight,viewWidth,viewHeight,blockLeft,blockTop,blockRight,blockBottom,fetchLeft,fetchTop,rangeLeft,rangeTop,rangeRight,rangeBottom);
				}else if("eastsouth".equals(direction)){
					result = jump("ES",(Spreadsheet)comp,sheetId,sheet,type,dpWidth,dpHeight,viewWidth,viewHeight,blockLeft,blockTop,blockRight,blockBottom,fetchLeft,fetchTop,rangeLeft,rangeTop,rangeRight,rangeBottom);
				}else{
					throw new UiException("Unknow direction:"+direction);
				}
				responseDataBlock("Jump", token, sheetId, result);
			} else if ("neighbor".equals(type)) {
				if("east".equals(direction)){
					
					int right = blockRight + fetchWidth ;//blockRight+ 1 + fetchWidth - 1;
					right = _mergeMatrix.getRightConnectedColumn(right,	blockTop, blockBottom);
					
					int size = right - blockRight;//right - (blockRight +1) +1

					String result = loadEast(sheet, type, blockLeft, blockTop, blockRight, blockBottom, size);
					responseDataBlock("East", token, sheetId, result);
				} else if ("south".equals(direction)) {
					
					int bottom = blockBottom + fetchHeight;
					
					//check right for new load south block
					int right = _mergeMatrix.getRightConnectedColumn(blockRight, blockTop, bottom);
					int left = _mergeMatrix.getLeftConnectedColumn(blockLeft, blockTop, bottom);

					if (right > blockRight) {
						String result = loadEast(sheet, type, blockLeft, blockTop, blockRight, blockBottom, right - blockRight);
						responseDataBlock("East", "", sheetId, result);
					}
					if (left < blockLeft) {
						String result = loadWest(sheet, type, blockLeft, blockTop, right, blockBottom, blockLeft - left);
						responseDataBlock("West", "", sheetId, result);
					}

					String result = loadSouth(sheet, type, left, blockTop, right, blockBottom, fetchHeight);
					responseDataBlock("South", token, sheetId, result);
				} else if ("west".equals(direction)) {
					
					int left = blockLeft - fetchWidth ;//blockLeft - 1 - fetchWidth + 1;
					left = _mergeMatrix.getLeftConnectedColumn(left,blockTop,blockBottom);
					int size = blockLeft - left ;//blockLeft -1 - left + 1;
					
					String result = loadWest(sheet, type, blockLeft, blockTop,	blockRight, blockBottom, size);
					responseDataBlock("West", token, sheetId, result);
				} else if("north".equals(direction)) {
					
					
					int top = blockTop - fetchHeight;
					
					//check right-left for new load north block
					int right = _mergeMatrix.getRightConnectedColumn(blockRight, top, blockBottom);
					int left = _mergeMatrix.getLeftConnectedColumn(blockLeft,top, blockBottom);
					
					if (right > blockRight) {
						String result = loadEast(sheet, type, blockLeft, blockTop, blockRight, blockBottom, right - blockRight);
						responseDataBlock("East", "", sheetId, result);
					}
					if (left < blockLeft) {
						String result = loadWest(sheet,type,blockLeft,blockTop,right,blockBottom,blockLeft - left);
						responseDataBlock("West", "", sheetId,result);
					}
					
					String result = loadNorth(sheet, type, left, blockTop, right, blockBottom, fetchHeight);
					responseDataBlock("North", token, sheetId, result);
				}
			} else if("visible".equals(type)) {
				loadForVisible((Spreadsheet) comp, sheetId, sheet, type, dpWidth, dpHeight, viewWidth, viewHeight, blockLeft, blockTop, blockRight, blockBottom, rangeLeft, rangeTop, rangeRight, rangeBottom);
				//always ack for call back
				String ack = ackResult();
				responseDataBlock(null, token,sheetId,ack);
			} else {
				//TODO use debug warning
				log.warning("unknow type:"+type);
			}
				
		} catch(Throwable x) {
			responseDataBlock("Error", "", sheetId, ackError(x.getMessage()));
			throw new UiException(x.getMessage(), x);
		}
		((SpreadsheetInCtrl) _ctrl).setLoadedRect(_lastleft, _lasttop,	_lastright, _lastbottom);
	}
	
	private void loadForVisible(Spreadsheet spreadsheet, String sheetId, Worksheet sheet, String type, int dpWidth,
			int dpHeight, int viewWidth, int viewHeight, int blockLeft,int blockTop,int blockRight, int blockBottom,
			int rangeLeft, int rangeTop, int rangeRight,int rangeBottom) {
		
		if (rangeRight > spreadsheet.getMaxcolumns() - 1) {
			rangeRight = spreadsheet.getMaxcolumns() - 1;
		}
		if (rangeBottom > spreadsheet.getMaxrows() - 1) {
			rangeBottom = spreadsheet.getMaxrows() - 1;
		}
		//calculate visible range , for merge range.
		int left = Math.min(rangeLeft, blockLeft);
		int top = Math.min(rangeTop, blockTop);
		int right = Math.max(rangeRight, blockRight);
		int bottom = Math.max(rangeBottom, blockBottom);

		if (right > blockRight) {
			
			right = _mergeMatrix.getRightConnectedColumn(right, top, bottom);

			int size = right - blockRight;
			String result = loadEast(sheet, type, blockLeft, blockTop, blockRight, blockBottom, size);
			responseDataBlock("East", "", sheetId, result);
			blockRight += size;
		}
		
		if (left < blockLeft) {

			left = _mergeMatrix.getLeftConnectedColumn(left, top, bottom);

			int size = blockLeft - left;
			String result = loadWest(sheet, type, blockLeft, blockTop, right, blockBottom, size);
			responseDataBlock("West", "", sheetId, result);
			blockLeft -= size;
		}

		if (bottom > blockBottom) {
			int size = bottom - blockBottom;
			String result = loadSouth(sheet, type, left, blockTop, right, blockBottom, size);
			responseDataBlock("South", "", sheetId, result);
			blockBottom += size;
		}

		if (top < blockTop) {
			int size = blockTop - top;
			String result = loadNorth(sheet, type, left, blockTop, right, bottom, size);
			responseDataBlock("North", "", sheetId, result);
			blockTop -= size;
		}
	}
	
	private String ackResult(){
		JSONObj jresult = new JSONObj();
		jresult.setData("type", "ack");
		return jresult.toString();
	}
	
	private String ackError(String message){
		JSONObj jresult = new JSONObj();
		jresult.setData("type","error");
		jresult.setData("message",message);
		return jresult.toString();
	}
	
	
	private String pruneResult(String dir, int reserve) {
		JSONObj jresult = new JSONObj();
		jresult.setData("type", "prune");
		jresult.setData("dir", dir);
		jresult.setData("reserve", reserve);
		return jresult.toString();
	}

	private String jumpResult(Worksheet sheet, int left, int top, int right, int bottom) {
		
		right = _mergeMatrix.getRightConnectedColumn(right,top,bottom);
		left = _mergeMatrix.getLeftConnectedColumn(left,top,bottom);

		int w = right - left + 1;
		int h = bottom - top + 1;
		
		//check merge range;
		JSONObj jresult = new JSONObj();
		List rd = new ArrayList();
		
		jresult.setData("type", "jump");
		jresult.setData("dir", "jump");
		jresult.setData("left", left);
		jresult.setData("top", top);
		jresult.setData("width", w);
		jresult.setData("height", h);
		jresult.setData("data", rd);
		
		// prepare header
		JSONObj jheader;
		List theaders = null;
		List lheaders = null;
		boolean filltheader = false;
		
		if (!_hidecolhead) {
			theaders = new ArrayList();
			jresult.setData("theader", theaders);
		}

		if (!_hiderowhead) {
			lheaders = new ArrayList();
			jresult.setData("lheader", lheaders);
		}
		
		JSONObj jrow, jcell;
		// append row
		int re = top + h;
		int ce = left + w;
		for (int i = top; i < re; i++) {
			jrow = new JSONObj();
			rd.add(jrow);
			prepareRowData(jrow, sheet, i);
			List cells = new ArrayList();
			jrow.setData("cells", cells);
			for (int j = left; j < ce; j++) {
				jcell = new JSONObj();
				Cell cell = Utils.getCell(sheet, i, j);
				cells.add(jcell);
				prepareCellData(jcell, sheet, i, j);
				if (!_hidecolhead && !filltheader) {
					jheader = new JSONObj();
					prepareTopHeaderData(jheader, j);
					theaders.add(jheader);
				}
			}
			filltheader = true;
			if (!_hiderowhead) {
				jheader = new JSONObj();
				prepareLeftHeaderData(jheader, i);
				lheaders.add(jheader);
			}
		}
		
		_lastleft = left;
		_lastright = right;
		_lasttop = top;
		_lastbottom = bottom;
		
		
		// prepare top frozen cell
		int fzr = _spreadsheet.getRowfreeze();
		if (fzr > -1) {
			JSONObj topFrozen = new JSONObj();
			jresult.setData("topfrozen", topFrozen);

			List tfrd = new ArrayList();

			topFrozen.setData("type", "jump");
			topFrozen.setData("dir", "jump");
			topFrozen.setData("left", left);
			topFrozen.setData("top", 0);
			topFrozen.setData("width", w);
			topFrozen.setData("height", fzr + 1);
			topFrozen.setData("data", tfrd);

			ce = left + w;
			for (int i = 0; i <= fzr; i++) {
				jrow = new JSONObj();
				tfrd.add(jrow);
				prepareRowData(jrow, sheet, i);
				List cells = new ArrayList();
				jrow.setData("cells", cells);
				for (int j = left; j < ce; j++) {
					jcell = new JSONObj(); // cell
					Cell cell = Utils.getCell(sheet, i, j);
					cells.add(jcell);
					prepareCellData(jcell, sheet, i, j);
				}
			}
		}
		
		
		//prepare left frozen cell
		int fzc = _spreadsheet.getColumnfreeze();
		if (fzc > -1) {
			JSONObj leftFrozen = new JSONObj();
			jresult.setData("leftfrozen", leftFrozen);
			
			List lfrd = new ArrayList();
			leftFrozen.setData("type", "jump");
			leftFrozen.setData("dir", "jump");
			leftFrozen.setData("left", 0);
			leftFrozen.setData("top", top);
			leftFrozen.setData("width", fzc + 1);
			leftFrozen.setData("height", h);
			leftFrozen.setData("data", lfrd);
			
			re = top + h;
			for (int i = top; i < re; i++) {
				jrow = new JSONObj();
				lfrd.add(jrow);
				prepareRowData(jrow, sheet, i);
				List cells = new ArrayList();
				jrow.setData("cells", cells);
				for (int j = 0; j <= fzc; j++) {
					jcell = new JSONObj(); // cell
					Cell cell = Utils.getCell(sheet, i, j);
					cells.add(jcell);
					prepareCellData(jcell, sheet, i, j);
				}
			}
		}
		return jresult.toString();
	}
	
	private String jump(String dir,Spreadsheet spreadsheet,String sheetId, Worksheet sheet, String type,
			int dpWidth, int dpHeight, int viewWidth, int viewHeight,
			int blockLeft, int blockTop, int blockRight, int blockBottom,
			int col, int row, 
			int rangeLeft, int rangeTop, int rangeRight,int rangeBottom) {
		
		int left;
		int right;
		int top;
		int bottom;
		
		
		if (dir.indexOf("E") >= 0) {
			right = col + 1;
			left = _colHelper.getCellIndex(_colHelper.getStartPixel(col) - viewWidth);
			// w = col - left + 2;//load more;

			if (right > spreadsheet.getMaxcolumns() - 1) {
				// w = spreadsheet.getMaxcolumn()-left;
				right = spreadsheet.getMaxcolumns() - 1;
			}
		} else if (dir.indexOf("W") >= 0) {
			left = col <= 0 ? 0 : col - 1;
			right = _colHelper.getCellIndex(_colHelper.getStartPixel(col)
					+ viewWidth);// end cell index

			if (right > spreadsheet.getMaxcolumns() - 1) {
				// w = spreadsheet.getMaxcolumn()-left;
				right = spreadsheet.getMaxcolumns() - 1;
			}
		} else {
			left = blockLeft;// rangeLeft;
			right = blockRight;// rangeRight;
		}
		
		if (dir.indexOf("S") >= 0) {
			bottom = row + 1;
			top = _rowHelper.getCellIndex(_rowHelper.getStartPixel(row)	- viewHeight);

			if (bottom > spreadsheet.getMaxrows() - 1) {
				bottom = spreadsheet.getMaxrows() - 1;
			}
		} else if (dir.indexOf("N") >= 0) {
			top = row <= 0 ? 0 : row - 1;
			bottom = _rowHelper.getCellIndex(_rowHelper.getStartPixel(row) + viewHeight);// end cell index

			if (bottom > spreadsheet.getMaxrows() - 1) {
				bottom = spreadsheet.getMaxrows() - 1;
			}
		} else {
			top = blockTop;// rangeTop;
			bottom = blockBottom;// rangeBottom;
		}
		
		return jumpResult(sheet,left,top,right,bottom);
	}
	
	private String loadEast(Worksheet sheet,String type, 
			int blockLeft,int blockTop,int blockRight, int blockBottom,
			int fetchWidth) {

		/**
		 * { "data":[ 
		 * 	{ "type":"row", "index":1, "cells":[ 
		 * 		{ "index":1,"txt":"", "format":""; }, 
		 * 		{ "index":1, "txt":"", "format":""; } 
		 * 	] }
		 * ] }
		 */

		JSONObj jresult = new JSONObj();
		List rd = new ArrayList();

		jresult.setData("type", "neighbor");
		jresult.setData("dir", "east");
		jresult.setData("width", fetchWidth);
		jresult.setData("height", blockBottom - blockTop + 1);
		jresult.setData("data", rd);
		JSONObj jrow, jcell;
		JSONObj jheader;
		List headers = null;
		if (!_hidecolhead) {
			headers = new ArrayList();
			jresult.setData("theader", headers);
		}
		boolean fillheader = false;
		
		//append row
		int cs = blockRight + 1;
		int ce = cs + fetchWidth;
		
		for (int i = blockTop; i <= blockBottom; i++) {
			jrow = new JSONObj();
			rd.add(jrow);
			prepareRowData(jrow, sheet, i);
			List cells = new ArrayList();

			jrow.setData("cells", cells);
			for (int j = cs; j < ce; j++) {
				jcell = new JSONObj(); // cell
				Cell cell = Utils.getCell(sheet, i, j);
				cells.add(jcell);
				prepareCellData(jcell, sheet, i, j);
				if (!_hidecolhead && !fillheader) {
					jheader = new JSONObj();
					prepareTopHeaderData(jheader, j);
					headers.add(jheader);
				}

			}
			fillheader = true;
		}
	
		_lastleft = blockLeft;
		_lastright = ce - 1;
		_lasttop = blockTop;
		_lastbottom = blockBottom;

		//process frozen row data
		int fzr = _spreadsheet.getRowfreeze();
		if (fzr > -1) {
			JSONObj topFrozen = new JSONObj();
			jresult.setData("topfrozen", topFrozen);

			List tfrd = new ArrayList();

			topFrozen.setData("type", "neighbor");
			topFrozen.setData("dir", "east");
			topFrozen.setData("width", fetchWidth);
			topFrozen.setData("height", fzr + 1);
			topFrozen.setData("data", tfrd);

			for (int i = 0; i <= fzr; i++) {
				jrow = new JSONObj();
				tfrd.add(jrow);
				prepareRowData(jrow, sheet, i);
				List cells = new ArrayList();
				jrow.setData("cells", cells);
				for (int j = cs; j < ce; j++) {
					jcell = new JSONObj(); // cell
					Cell cell = Utils.getCell(sheet, i, j);
					cells.add(jcell);
					prepareCellData(jcell, sheet, i, j);
				}
			}
		}
		return jresult.toString();
	}
	
	private String loadWest(Worksheet sheet,String type,
			int blockLeft,int blockTop,int blockRight, int blockBottom,
			int fetchWidth) {
		
		JSONObj jresult = new JSONObj();
		List rd = new ArrayList();

		jresult.setData("type", "neighbor");
		jresult.setData("dir", "west");
		jresult.setData("width", fetchWidth);// increased cell size
		jresult.setData("height", blockBottom - blockTop + 1);// increased cell size
		jresult.setData("data", rd);
		JSONObj jrow, jcell;
		JSONObj jheader;
		List headers = null;
		if (!_hidecolhead) {
			headers = new ArrayList();
			jresult.setData("theader", headers);
		}
		boolean fillheader = false;
		
		// append row
		int cs = blockLeft - 1;
		int ce = cs - fetchWidth;
		for (int i = blockTop; i <= blockBottom; i++) {
			jrow = new JSONObj();
			rd.add(jrow);
			prepareRowData(jrow, sheet, i);
			List cells = new ArrayList();
			jrow.setData("cells", cells);
			for (int j = cs; j > ce; j--) {
				jcell = new JSONObj();
				Cell cell = Utils.getCell(sheet, i, j);
				cells.add(jcell);
				prepareCellData(jcell, sheet, i, j);

				if (!_hidecolhead && !fillheader) {
					jheader = new JSONObj();
					prepareTopHeaderData(jheader, j);
					headers.add(jheader);
				}
			}
			fillheader = true;
		}
		
		_lastleft = ce+1;
		_lastright = blockRight;
		_lasttop = blockTop;
		_lastbottom = blockBottom;
		
		// process frozen row data
		int fzr = _spreadsheet.getRowfreeze();
		if (fzr > -1) {
			JSONObj topFrozen = new JSONObj();
			jresult.setData("topfrozen", topFrozen);

			List tfrd = new ArrayList();

			topFrozen.setData("type", "neighbor");
			topFrozen.setData("dir", "west");
			topFrozen.setData("width", fetchWidth);
			topFrozen.setData("height", fzr + 1);
			topFrozen.setData("data", tfrd);

			for (int i = 0; i <= fzr; i++) {
				jrow = new JSONObj();
				tfrd.add(jrow);
				prepareRowData(jrow, sheet, i);
				List cells = new ArrayList();
				jrow.setData("cells", cells);
				for (int j = cs; j > ce; j--) {
					jcell = new JSONObj(); // cell
					Cell cell = Utils.getCell(sheet, i, j);
					cells.add(jcell);
					prepareCellData(jcell, sheet, i, j);
				}
			}
		}
		
		return jresult.toString();
	}
	
	private String loadSouth(Worksheet sheet,String type, 
			int blockLeft,int blockTop,int blockRight, int blockBottom,
			int fetchHeight) {
		
		JSONObj jresult = new JSONObj();
		List rd = new ArrayList();

		jresult.setData("type", "neighbor");
		jresult.setData("dir", "south");
		jresult.setData("width", blockRight - blockLeft + 1);
		jresult.setData("height", fetchHeight);
		jresult.setData("data", rd);
		JSONObj jrow, jcell;
		JSONObj jheader;
		List headers = null;
		if (!_hiderowhead) {
			headers = new ArrayList();
			jresult.setData("lheader", headers);
		}

		int rs = blockBottom + 1;
		int re = rs + fetchHeight;

		for (int i = rs; i < re; i++) {
			jrow = new JSONObj();
			rd.add(jrow);
			prepareRowData(jrow, sheet, i);
			List cells = new ArrayList();
			jrow.setData("cells", cells);
			for (int j = blockLeft; j <= blockRight; j++) {
				jcell = new JSONObj();
				Cell cell = Utils.getCell(sheet, i, j);
				cells.add(jcell);
				prepareCellData(jcell, sheet, i, j);
			}
			if (!_hiderowhead) {
				jheader = new JSONObj();
				prepareLeftHeaderData(jheader, i);
				headers.add(jheader);
			}
		}
		_lastleft = blockLeft;
		_lastright = blockRight;
		_lasttop = blockTop;
		_lastbottom = re-1;
		
		// process frozen left
		int fzc = _spreadsheet.getColumnfreeze();
		if (fzc > -1) {
			JSONObj leftFrozen = new JSONObj();
			jresult.setData("leftfrozen", leftFrozen);

			List lfrd = new ArrayList();

			leftFrozen.setData("type", "neighbor");
			leftFrozen.setData("dir", "south");
			leftFrozen.setData("width", fzc + 1);
			leftFrozen.setData("height", fetchHeight);
			leftFrozen.setData("data", lfrd);

			rs = blockBottom + 1;
			re = rs + fetchHeight;

			for (int i = rs; i < re; i++) {
				jrow = new JSONObj();
				lfrd.add(jrow);
				prepareRowData(jrow, sheet, i);
				List cells = new ArrayList();
				jrow.setData("cells", cells);
				for (int j = 0; j <= fzc; j++) {
					jcell = new JSONObj(); // cell
					Cell cell = Utils.getCell(sheet, i, j);
					cells.add(jcell);
					prepareCellData(jcell, sheet, i, j);
				}
			}
		}

		return jresult.toString();
	}
	private String loadNorth(Worksheet sheet,String type, 
			int blockLeft,int blockTop,int blockRight, int blockBottom,
			int fetchHeight) {

		JSONObj jresult = new JSONObj();
		List rd = new ArrayList();
		
		jresult.setData("type", "neighbor");
		jresult.setData("dir", "north");
		jresult.setData("width", blockRight - blockLeft + 1);
		jresult.setData("height", fetchHeight);
		jresult.setData("data", rd);
		JSONObj jrow, jcell;
		JSONObj jheader;
		List headers = null;
		if (!_hiderowhead) {
			headers = new ArrayList();
			jresult.setData("lheader", headers);
		}
		// append row
		int rs = blockTop - 1;
		int re = rs - fetchHeight;
		for (int i = rs; i > re; i--) {
			jrow = new JSONObj();
			rd.add(jrow);
			prepareRowData(jrow, sheet, i);
			List cells = new ArrayList();
			jrow.setData("cells", cells);
			for (int j = blockLeft; j <= blockRight; j++) {
				jcell = new JSONObj();
				Cell cell = Utils.getCell(sheet, i, j);
				cells.add(jcell);
				prepareCellData(jcell, sheet, i, j);
			}
			if (!_hiderowhead) {
				jheader = new JSONObj();
				prepareLeftHeaderData(jheader, i);
				headers.add(jheader);
			}
		}
		_lastleft = blockLeft;
		_lastright = blockRight;
		_lasttop = re + 1;
		_lastbottom = blockBottom;
		
		
		// process frozen left
		int frc = _spreadsheet.getColumnfreeze();
		if (frc > -1) {
			JSONObj leftFrozen = new JSONObj();
			jresult.setData("leftfrozen", leftFrozen);

			List lfrd = new ArrayList();

			leftFrozen.setData("type", "neighbor");
			leftFrozen.setData("dir", "north");
			leftFrozen.setData("width", frc + 1);
			leftFrozen.setData("height", fetchHeight);
			leftFrozen.setData("data", lfrd);

			rs = blockTop - 1;
			re = rs - fetchHeight;

			for (int i = rs; i > re; i--) {
				jrow = new JSONObj();
				lfrd.add(jrow);
				prepareRowData(jrow, sheet, i);
				List cells = new ArrayList();
				jrow.setData("cells", cells);
				for (int j = 0; j <= frc; j++) {
					jcell = new JSONObj(); // cell
					Cell cell = Utils.getCell(sheet, i, j);
					cells.add(jcell);
					prepareCellData(jcell, sheet, i, j);
				}
			}
		}
		
		return jresult.toString();
	}
	
	private void prepareTopHeaderData(JSONObj jheader, int col) {
		jheader.setData("ix", col);
		String name = (String) _spreadsheet.getColumntitle(col);
		jheader.setData("nm", name);
		final HeaderPositionInfo info = _colHelper.getInfo(col);
		if (info != null) {
			jheader.setData("zsw", info.id);
			if (info.hidden) {
				jheader.setData("hn", true);
			}
		}
	}
	
	private void prepareLeftHeaderData(JSONObj jheader, int row) {
		jheader.setData("ix", row);
		String name = (String) _spreadsheet.getRowtitle(row);
		jheader.setData("nm", name);
		final HeaderPositionInfo info = _rowHelper.getInfo(row);
		if (info != null) {
			jheader.setData("zsh", info.id);
			if (info.hidden) {
				jheader.setData("hn", true);
			}
		}
	}
	
	private void prepareRowData(JSONObj jrow, Worksheet sheet,int row) {
		jrow.setData("ix", row);
		HeaderPositionInfo info = _rowHelper.getInfo(row);
		if (info != null)
			jrow.setData("zsh", info.id);
	}
	
	private void prepareCellData(JSONObj jcell, Worksheet sheet,int row,int col) {

		jcell.setData("ix", col);// index
		
		Cell cell = Utils.getCell(sheet, row, col); 
		
		if (cell != null) {
			CellStyle style = cell.getCellStyle();
			
			boolean wrap = false;
			
			CellFormatHelper cfh = new CellFormatHelper(sheet, row, col, _mergeMatrix);
			String st = cfh.getHtmlStyle();
			String ist = cfh.getInnerHtmlStyle();
			if (st != null && !"".equals(st)) {
				jcell.setData("st", st);// style of text cell.
			}
			if (ist != null && !"".equals(ist)) {
				jcell.setData("ist", ist);// inner style of text cell
			}
			if (style != null && style.getWrapText()) {
				wrap = true;
				jcell.setData("wrap", true);// warp
			}
			if (cfh.hasRightBorder()) {
				jcell.setData("rbo", true);// right border, when processing text overflow, must take care this.
			}

			final FormatText ft = Utils.getFormatText(cell);
			RichTextString rstr = ft != null && ft.isRichTextString()? ft.getRichTextString() : null;
			String text = rstr == null ? ft != null ? Utils.escapeCellText(ft.getCellFormatResult().text, wrap, true) : "" : Utils.formatRichTextString(sheet, rstr, wrap);
			Hyperlink hlink = Utils.getHyperlink(cell);
			if (hlink != null) {
				text = Utils.formatHyperlink(sheet, hlink, text, wrap);
			}
			jcell.setData("txt", text);
			jcell.setData("edit", Utils.getEditText(cell));

			int textHAlign = BookHelper.getRealAlignment(cell);
			switch(textHAlign) {
			case CellStyle.ALIGN_CENTER:
			case CellStyle.ALIGN_CENTER_SELECTION:
				jcell.setData("hal", "c");
				break;
			case CellStyle.ALIGN_RIGHT:
				jcell.setData("hal", "r");
				break;
			}
		} else {
			jcell.setData("txt", "");
			CellFormatHelper cfh = new CellFormatHelper(sheet, row, col, _ctrl.getMergeMatrixHelper(sheet));
			String st = cfh.getHtmlStyle();
			String ist = cfh.getInnerHtmlStyle();
			if (st != null && !"".equals(st))
				jcell.setData("st", st);// style of text cell.

			if (ist != null && !"".equals(ist))
				jcell.setData("ist", ist);// inner style of text cell

			if (cfh.hasRightBorder())
				jcell.setData("rbo", true);// format
		}
		
		MergedRect rect = _mergeMatrix.getMergeRange(row, col);
		if ((rect != null)) {
			jcell.setData("merr", rect.getRight());// right
			jcell.setData("merid", rect.getId());// right
			jcell.setData("merl", rect.getLeft());// left
		}

		HeaderPositionInfo info = _colHelper.getInfo(col);
		if (info != null) {
			jcell.setData("zsw", info.id);
		}
		
	}
}