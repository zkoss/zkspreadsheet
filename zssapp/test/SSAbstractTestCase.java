import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;

import org.junit.Test;
import org.zkoss.ztl.JQuery;
import org.zkoss.ztl.ZKClientTestCase;

import com.thoughtworks.selenium.Selenium;

/**
 * An abstract test case class for ZSS. You can just write your test action in your case which
 * extends this class.
 * @author Phoenix
 *
 */
public abstract class SSAbstractTestCase extends ZKClientTestCase {
    public SSAbstractTestCase() {
        target = Utils.getTarget();
        browsers = getBrowsers(Utils.getBrowsers());
        _timeout = Utils.getTimeout();
    }
    
    @Test(expected = AssertionError.class)
    public final void testCase() {
        for (Selenium browser : browsers) {
            try {
                start(browser);
                windowFocus();
                waitResponse();
                windowMaximize();
                sleep(2000L);
                waitResponse();
                executeTest();
            }catch(Exception e){
            	e.printStackTrace();
            } finally {
                stop();
            }
        }
    }
    
    /**
     * 
     * get style may have different result.
     * 
     * Get the style of certain cell
     * @param col
     * @param row
     * @return
     */
    public String getCellStyle(int col, int row){
    	if (isIE6() || isIE7()) {
    		return getSpecifiedCell(col, row).attr("style");
    	}
    	return jq("div.zscell[z\\\\.c=\"" + col + "\"][z\\\\.r=\"" + row + "\"]").first().attr("style")
    	+jq("div.zscell[z\\\\.c=\"" + col + "\"][z\\\\.r=\"" + row + "\"] div").first().attr("style");
    }
    
    public String getCellBackgroundColor(int col, int row) {
    	String style = getSpecifiedCellOuter(col, row).attr("style");
    	int startIdx = Math.max(style.indexOf("background-color:"), style.indexOf("BACKGROUND-COLOR:")) ;
    	if (startIdx < 0)
    		return "";
    	int endIdx = style.indexOf(";", startIdx);
    	endIdx = endIdx > 0 ? endIdx : style.length();
    	return style.substring(startIdx + "background-color:".length(), endIdx).trim();
    }
    
    public String getCellFontFamily(int col, int row) {
    	String font = getSpecifiedCell(col, row).css("font-family");
    	return font != null ? font.replaceAll("'", "") : "";
    }
    
    public String getCellFontColor(int col, int row) {
    	String color = getSpecifiedCell(col, row).css("color");
    	return color;
    }
    
    public final static int BORDER_TOP = 0;
    public final static int BORDER_RIGHT = 1;
    public final static int BORDER_BOTTOM = 2;
    public final static int BORDER_LEFT = 3;

    public String getCellBorderColor(int col, int row, int borderType) {
    	JQuery cell = getSpecifiedCellOuter(col, row);
    	String style = cell.attr("style");
    	
    	int startIdx = -1;
    	int borderLength = 0;
    	switch (borderType) {
    	case BORDER_TOP:
    		borderLength = "border-top:".length();
    		startIdx = Math.max(style.indexOf("border-top:"), style.indexOf("BORDER-TOP:"));
    		break;
    	case BORDER_RIGHT:
    		borderLength = "border-right:".length();
    		startIdx = Math.max(style.indexOf("border-right:"), style.indexOf("BORDER-RIGHT:"));
    		break;
    	case BORDER_BOTTOM:
    		borderLength = "border-bottom:".length();
    		startIdx = Math.max(style.indexOf("border-bottom:"), style.indexOf("BORDER-BOTTOM:"));
    		break;
    	case BORDER_LEFT:
    		borderLength = "border-right:".length();
    		startIdx = Math.max(style.indexOf("border-right:"), style.indexOf("BORDER-RIGHT:"));
    		break;
    	}
    	
		if (startIdx < 0)
			return "";
    	int endIdx = style.indexOf(";", startIdx);
    	endIdx = endIdx > 0 ? endIdx : style.length();
    	return style.substring(startIdx + borderLength, endIdx).trim();
    }
    
    public String getCellFontAlign(int col, int row) {
    	return getSpecifiedCell(col, row).css("text-align");
    }
    

//    /**
//     * Set cell font color by top toolbar button
//     * 
//     * @param lCol
//     * @param tRow
//     * @param rCol
//     * @param bRow
//     * @param nthColorPalette
//     * @return
//     */
//    public String setCellFontColorByToolbarbutton(int lCol, int tRow, int rCol, int bRow, int nthColorPalette) {
//    	click(jq("$fontCtrlPanel $fontColorBtn"));
//    	click(jq());
//    }
    public String setCellFontColorByToolbarbutton(int col, int row, int nthColorPalette) {
    	return setCellFontColorByToolbarbutton(col, row, col, row, nthColorPalette);
    }
    
    public String setCellFontColorByToolbarbutton(int lCol, int tRow, int rCol, int bRow, int nthColorPalette) {
    	selectCells(lCol, tRow, rCol, bRow);
    	click(jq("$fontCtrlPanel $fontColorBtn"));
        JQuery color = jq(".z-colorpalette:visible div.z-colorpalette-colorbox:nth-child(" + nthColorPalette + ")");
        String selectedColor = color.first().text();
    	mouseOver(color);
    	click(color);
    	waitResponse();
    	return selectedColor;
    }
    
    public String setCellBackgroundColorByToolbarbutton(int col, int row, int nthColorPalette) {
    	return setCellBackgroundColorByToolbarbutton(col, row, col, row, nthColorPalette);
    }
    
    public String setCellBackgroundColorByToolbarbutton(int lCol, int tRow, int rCol, int bRow, int nthColorPalette) {
    	selectCells(lCol, tRow, rCol, bRow);
    	click(jq("$fontCtrlPanel $cellColorBtn"));
        JQuery color = jq(".z-colorpalette:visible div.z-colorpalette-colorbox:nth-child(" + nthColorPalette + ")");
        String selectedColor = color.first().text();
    	mouseOver(color);
    	click(color);
    	waitResponse();
    	return selectedColor;
    }
    
    /**
     * Sets cell font color by fast toolbar button
     */
    public String setCellFontColorByFastToolbarbutton(int lCol, int tRow, int rCol, int bRow, int nthColorPalette) {
    	rightClickCells(lCol, tRow, rCol, bRow);
    	click("$cellContext $fontCtrlPanel:visible $fontColorBtn");
    	
        JQuery color = jq(".z-colorpalette:visible div.z-colorpalette-colorbox:nth-child(" + nthColorPalette + ")");
        String selectedColor = color.first().text();
    	mouseOver(color);
    	click(color);
    	waitResponse();
    	return selectedColor;
    }
    
    public String setCellBackgroundColorByFastToolbarbutton(int lCol, int tRow, int rCol, int bRow, int nthColorPalette) {
    	rightClickCells(lCol, tRow, rCol, bRow);
    	click("$cellContext $fontCtrlPanel:visible $cellColorBtn");
    	
        JQuery color = jq(".z-colorpalette:visible div.z-colorpalette-colorbox:nth-child(" + nthColorPalette + ")");
        String selectedColor = color.first().text();
    	mouseOver(color);
    	click(color);
    	waitResponse();
    	return selectedColor;
    }
    
    public String setCellBackgroundColorByFastToolbarbutton(int col, int row, int nthColorPalette) {
    	return setCellBackgroundColorByFastToolbarbutton(col, row, col, row, nthColorPalette);
    }
    
    /**
     * Sets cell font color
     * 
     * @param col
     * @param row
     * @param nthColorPalette
     * @return
     */
    public String setCellFontColorByFastToolbarbutton(int col, int row, int nthColorPalette) {
    	return setCellFontColorByFastToolbarbutton(col, row, col, row, nthColorPalette);
    }
    
    public static final String CELL_WITHOUT_STYLE = "rgba(0, 0, 0, 0):Arial:left";
    public static final String CELL_WITHOUT_STYLE2 = "transparent:Arial:left";
    
    /**
     * 
     * Get some style related property of a cell. Including background-color, font-family, text-align
     * "rgba(0, 0, 0, 0):Arial:left" or "transparent:Arial:left" is the value for cleared style
     * @param col
     * @param row
     * @return
     */
    public String getCellCompositeStyle(int col, int row){
    	String result ="";
    	if (isIE6() || isIE7()) {
    		result += getCellBackgroundColor(col, row);
    		result += (":" + getCellFontFamily(col, row));
    		result += (":" + getCellFontAlign(col, row));
    		return result;
    	}
    	
    	result += jq("div.zscell[z\\\\.c=\"" + col + "\"][z\\\\.r=\"" + row + "\"]").css("background-color");
    	result += ":"+jq("div.zscell[z\\\\.c=\"" + col + "\"][z\\\\.r=\"" + row + "\"] div").css("font-family");
    	result += ":"+jq("div.zscell[z\\\\.c=\"" + col + "\"][z\\\\.r=\"" + row + "\"] div").css("text-align");
    	return result;
    }
    
//    public boolean isDefaultCellStyle(int row, int col) {
//    	
//    }
//    
//    public boolean isSameCellStyle() {
//    	
//    }
    
    /**
     * 
     * @param col - Base on 0
     * @param row - Base on 0
     * @return A JQuery object of cell. (Inner div)
     */
    public JQuery getSpecifiedCell(int col, int row) {
    	if (isIE6() || isIE7()) {
    		JQuery r = getRow(row);
    		JQuery cell = r.children().eq(col);
    		return cell.children().first();
    	}
    	return jq("div.zscell[z\\\\.c=\"" + col + "\"][z\\\\.r=\"" + row + "\"] div").first();	
    }
    
    public String getCellTextColor(int col, int row) {
    	String cellStyle = getCellStyle(col, row);
    	cellStyle = cellStyle.replaceAll(" ", "");
    	int startIdx = Math.max(cellStyle.indexOf(";color:"), cellStyle.indexOf(";COLOR:"));
    	return startIdx >=0 ? cellStyle.substring(startIdx + 1, cellStyle.indexOf(";", startIdx + ";color:".length())) : "";
    }
    
//    public String getCellBackground(int col, int row) {
//    	
//    }

    private boolean isFF3(){
    	String ffVer = String.valueOf(getEval("zk.ff"));
    	return ffVer.startsWith("3");
    }
    
    private boolean isIE() {
    	String ieVer = String.valueOf(getEval("zk.ie"));
    	return ieVer != null && ieVer.length() > 0;
    }
    
    private boolean isIE7() {
    	return isIE() && isIE(7);
    }
    private boolean isIE6() {
    	return isIE() && isIE(6);
    }
    
    private boolean isIE(int version) {
    	if (!isIE())
    		return false;
    	Integer ver = null;
    	try {
    		ver = Integer.valueOf(getEval("zk.ie"));
    	} catch (NumberFormatException ex) {
    		return false;
    	}
    	return ver != null && ver == version;
    }
    
    /**
     * 
     * @param col - Base on 0
     * @param row - Base on 0
     * @return A JQuery object of cell Outer div.
     */    
    public JQuery getSpecifiedCellOuter(int col, int row) {
    	if (isIE6() || isIE7())
    		return getSpecifiedCell(col, row).parent();
        return jq("div.zscell[z\\\\.c=\"" + col + "\"][z\\\\.r=\"" + row + "\"]");
    }

    /**
     * 
     * @param col - Base on 0
     * @return A JQuery object of column header.
     */
    public JQuery getColumnHeader(int col) {
    	if (isIE6() || isIE7()) {
    		JQuery cols =  jq("div.zstopcell");
    		return cols.eq(col);
    	}
        return jq("div.zstopcell[z\\\\.c=\"" + col + "\"] div");
    }
    
    /**
     * 
     * @param row - Base on 0
     * @return A JQuery object of row header.
     */
    public JQuery getRowHeader(int row) {
    	if (isIE6() || isIE7()) {
    		JQuery rows = jq("div.zsleftcell");
    		return rows.eq(row);
    	}
       return jq("div.zsleftcell[z\\\\.r=\"" + row + "\"] div"); 
    }
    
    public String getCellText(int col, int row) {
    	if (isFF3() || isIE()) {
    		JQuery cell = getSpecifiedCell(col, row).children().children();
    		Iterator<JQuery> i = cell.iterator();
    		while (i.hasNext()) {
    			JQuery j = i.next();
    			if (!j.isVisible())
    				continue;
    			String str = j.text();
    			if (str != null && str.length() > 0) {
    				return str;
    			}
    		}
    		return "";
    	}
    	return getSpecifiedCell(col, row).text();
    }
    
    public String getCellText(JQuery cellLocator) {
    	if (isFF3() || isIE()) {
    		JQuery cell = cellLocator.children().children();
    		Iterator<JQuery> i = cell.iterator();
    		while (i.hasNext()) {
    			JQuery j = i.next();
    			if (!j.isVisible())
    				continue;
    			String str = j.text();
    			if (str != null && str.length() > 0) {
    				return str;
    			}
    		}
    		return "";
    	}
        return cellLocator.text();
    }
    
    /**
     * {@link Deprecated}
     * 
     * @param cellLocator
     */
    public void clickCell(JQuery cellLocator) {
    	//works for FF3,
    	mouseMoveAt(cellLocator, "4,4");
    	waitResponse();
        mouseDownAt(cellLocator, "4,4");
        waitResponse();
        mouseUpAt(cellLocator, "4,4");
        waitResponse();
        
//        mouseDown(cellLocator);
//        waitResponse();
//        mouseUp(cellLocator);
//        waitResponse();
    }
    
    public void dragFill(int srcCol, int srcRow, int targetCol, int targetRow) {
    	focusOnCell(srcCol, srcRow);
    	mouseOver(jq(".zsseldot"));
    	mouseDownAt(jq(".zsseldot"), ""+ 1 + "," + 1);
    	mouseMoveAt(getSpecifiedCell(targetCol, targetRow), "2,2");
    	mouseUpAt(getSpecifiedCell(targetCol, targetRow), "2,2");
    	waitResponse();
    }
    
    public void rightClickCell(JQuery cellLocator) {
        contextMenuAt(cellLocator, "2,2");
    }
    
    /**
     *  
     * @param column - Base on 0
     */
    public void rightClickColumnHeader(int column){    	
    	mouseDown(getColumnHeader(column));				
		waitResponse();
		mouseUp(getColumnHeader(column));
		waitResponse();
		mouseDown(getColumnHeader(column));
		waitResponse();
		mouseUp(getColumnHeader(column));
		waitResponse();
		contextMenuAt(getColumnHeader(column),"2,2");
		waitResponse();
    }
    
    /**
     * 
     * @param row : 0-based
     */
    public void rightClickRowHeader(int row){
		mouseDown(getRowHeader(row));
		waitResponse();
		mouseUp(getRowHeader(row));
		waitResponse();
		mouseDown(getRowHeader(row));
		waitResponse();
		mouseUp(getRowHeader(row));
		waitResponse();
		contextMenuAt(getRowHeader(row),"2,2");
		waitResponse();
    }
    
    /**
     * 
     * @param column: 0-based
     * @param row: 0-based
     */
    public void rightClickCell(int column, int row){
		mouseDownAt(getSpecifiedCell(column, row),"4,4");
		waitResponse();
		mouseUpAt(getSpecifiedCell(column, row),"4,4");
		waitResponse();
	
		mouseDownAt(getSpecifiedCell(column, row),"4,4");
		waitResponse();
		mouseUpAt(getSpecifiedCell(column, row),"4,4");
		waitResponse();

		contextMenuAt(getSpecifiedCell(column, row), "4,4");
		waitResponse();
    }

    /**
     * Selects group of cells
     * @param left: column number of left top cell, 0-based
     * @param top: row number of left top cell, 0-based
     * @param right: column number of right bottom cell, 0-based
     * @param bottom: row number of right bottom cell, 0-based
     */
    public void selectCells(int left, int top, int right, int bottom) {
    	//It's necessary to duplicate mouseDownAt target cell,
    	//because the first time may not really focused on target cell,
    	//but focus on cell A1
		mouseDownAt(getSpecifiedCell(left, top),"4,4");
		waitResponse();
		mouseUpAt(getSpecifiedCell(left, top),"4,4");
		waitResponse();
	
		mouseDownAt(getSpecifiedCell(left, top),"4,4");
		waitResponse();
		mouseMoveAt(getSpecifiedCell(right, bottom),"4,4");
		waitResponse();
		mouseUpAt(getSpecifiedCell(right, bottom),"4,4");
		waitResponse();
    }
    
    public JQuery focusOnCell(int left, int top) {
    	selectCells(left, top, left, top);
    	return getSpecifiedCell(left, top);
    }
    
    public boolean isFocusOnCell(int left, int top) {
    	return isFocusOnCell(getSpecifiedCell(left, top));
    }
    
    public boolean isFocusOnCell(JQuery cell) {
    	boolean isDisplayFocus = jq("div.zsscroll .zsdata .zsselect").isVisible();
    	int[] cellOffset = cell.zk().revisedOffset();
    	int[] focusOffset = jq("div.zsselect").zk().revisedOffset();
    	
    	//tested on chrome only
    	return isDisplayFocus && 
    		(cellOffset[0] - 4 == focusOffset[0]) &&
    		(cellOffset[1] - 2 == focusOffset[1]);
    }
    
    public boolean isFocusVisible() {
    	return jq("div.zsscroll .zsdata .zsselect").isVisible();
    }
    
    public boolean isHighlighVisible() {
    	return jq("div.zsscroll .zsdata .zshighlight").isVisible();
    }
    
    /**
     * 
     * @param left: column number of left top cell, 0-based
     * @param top: row number of left top cell, 0-based
     * @param right: column number of right bottom cell, 0-based
     * @param bottom: row number of right bottom cell, 0-based
     */
    public void rightClickCells(int left, int top, int right, int bottom){
    	selectCells(left, top, right, bottom);
    	
    	//offset "2,2" is too small, therefore, I tested with "10,10" and it works
		contextMenuAt(jq("div.zsselect .zsselecti"), "10,10"); 
		waitResponse();
    }

    /**
     * 
     * @param start: column number, 0-based
     * @param end: column number, 0-based
     */
    public void selectColumns(int start, int end){
		mouseDownAt(getColumnHeader(start),"2,2");
		waitResponse();
		mouseMoveAt(getColumnHeader(end),"2,2");
		waitResponse();
		mouseUpAt(getColumnHeader(end),"2,2");
    }

    /**
     * 
     * @param start: row number, 0-based
     * @param end: row number, 0-based
     */
    public void selectRows(int start, int end){
		mouseDownAt(getRowHeader(start),"2,2");
		waitResponse();
		mouseMoveAt(getRowHeader(end),"2,2");
		waitResponse();
		mouseUpAt(getRowHeader(end),"2,2");
    }

    /**
     * Simulate keys like CTRL+C, CTRL+X
     * @param keycode
     */
    public void pressCtrlWithChar(String keycode) {
		keyDownNative(CTRL);
		waitResponse();
		keyDownNative(keycode);
		waitResponse();
		keyUpNative(keycode);
		waitResponse();
		keyUpNative(CTRL);
		waitResponse();
    }
    
    /**
     * Returns number string, remove all others
     * @param val
     * @return
     */
    public String numberOnly(String val) {
    	return val.replaceAll("[^0-9]", "");
    }
    
    public boolean containsIgnoreCase(String source, String target) {
    	return source.toUpperCase().contains(target.toUpperCase());
    }
    
    public Map<String, String> getCellStyleMap(JQuery cellLocator) {
        Map<String, String> styleMap = new HashMap<String, String>();
        String[] style1 = null;
        String cellParentStyle = cellLocator.parent().attr("style");
        String cellStyle = cellLocator.attr("style");
        
        if (cellParentStyle.indexOf(";") > -1) {
            style1 = cellParentStyle.split(";");
            fillMap(styleMap, style1);
        }
        
        if (cellStyle.indexOf(";") > -1) {
            style1 = cellStyle.split(";");
            fillMap(styleMap, style1);
        }
        
        return styleMap;
    }
    
    private void fillMap(Map<String, String> map, String[] style) {
        for (String s : style) {
            int separatorIndex = s.indexOf(":");
            map.put(s.substring(0, separatorIndex).trim(), s.substring(separatorIndex + 1).trim());
        }
    }
    
    /**
     * Actually executing test.<br />
     * Note: You also have to implement your validation in this method.
     */
    protected abstract void executeTest();
    
    /**
     * Returns whether widget is exist and visible 
     * @param selector
     * @return
     */
    public boolean isWidgetVisible(String selector) {
    	return widget(jq(selector)).exists() && jq(selector).isVisible();
    }
    
    /**
     * Click dropdown button at nth-menu
     * @param selector
     */
    public void clickDropdownButtonMenu(String buttonSelector, int nthIndex) {
    	JQuery jq = jq(buttonSelector);
    	mouseOver(jq);
    	clickAt(jq, "30, 0");
    	click(jq(".z-menu-popup:visible .z-menu-item:eq(" + nthIndex + ")"));
    	waitResponse();
    }
    
    /**
     * Click dropdown button at menu
     * @param selector
     * @param string
     */
    public void clickDropdownButtonMenu(String buttonSelector, String menuLabel) {
    	JQuery jq = jq(buttonSelector);
    	mouseOver(jq);
    	clickAt(jq, "30, 0");
    	JQuery menus = jq(".z-menu-popup:visible .z-menu-item .z-menu-item-cnt");
    	int size = menus.length();
    	for (int i = 0; i < size; i++) {
    		JQuery menu = jq(".z-menu-popup:visible .z-menu-item .z-menu-item-cnt:eq(" + i + ")");
    		if (menu.text().indexOf(menuLabel) >= 0) {
    			click(menu.parent());
    			return;
    		}
    	}
    }
    
    /**
     * Returns row JQuery elements
     * @param row - Base on 0
     * @return A JQuery object of row.
     */
    public JQuery getRow(int row) {
    	if (isIE6() || isIE7()) {
    		JQuery r = jq(".zsscroll .zsblock").children().eq(row);
    		if (r.length() >= 1)
    			return r;
    	}
    	return jq(".zsscroll div.zsrow[z\\\\.r=\"" + row + "\"]"); //IE7 does not work
    }
    
//    public void setFontColor(String hex) {
//    	click(jq("$fontColorBtn"));
//    	waitResponse();
//        JQuery colorTextbox = jq(".z-colorpalette-hex-inp:visible");
//        type(colorTextbox, hex);
//        waitResponse();
//        keyPressEnter(colorTextbox);
//    }
//    
//    public void setCellBackgroundColor(String hex) {
//    	click(jq("$cellColorBtn"));
//    	waitResponse();
//        JQuery colorTextbox = jq(".z-colorpalette-hex-inp:visible");
//        type(colorTextbox, hex);
//        waitResponse();
//        keyPressEnter(colorTextbox);
//    }
    
    /**
     *  Set cell plain text as input by the end user.
     */
    public void setCellEditText(int col, int row, String txt) {
    	JQuery cell = getSpecifiedCell(col, row);
    	clickCell(cell);
    	typeKeys(cell, txt);
    	keyPressEnter(jq(".zsedit"));
    }
    
    final int RIGHT = 0;
    final int DOWN = 1;
    final int LEFT = 2;
    final int UP = 3;
    public void setRangeEditText(int col, int row, int orientation, String[] data) {
    	//ensure focus
    	JQuery cell = getSpecifiedCell(col, row);
    	clickCell(cell);
    	
    	int topRow = row;
    	int btmRow = row;
    	int leftCol = col;
    	int rightCol = col;
    	
    	final int size = data.length;
    	switch (orientation) {
    	case RIGHT:
    		rightCol += size;
    		break;
    	case DOWN:
    		btmRow += size;
    		break;
    	case LEFT:
    		leftCol -= size;
    		leftCol++;
    		break;
    	case UP:
    		topRow -= size;
    		topRow++;
    		break;
    	}
    	
    	int iter = 0;
    	for (int r = topRow; r <= btmRow; r++) {
    		for (int c = leftCol; c <= rightCol; c++) {
    			if (r <= 0 || c <= 0)
    				continue;
    			if (iter < size) {
    				setCellEditText(c, r, data[iter++]);
    			}
    		}
    	}
    }
}
