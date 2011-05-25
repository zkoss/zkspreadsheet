import java.util.HashMap;
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
    	return jq("div.zscell[z\\\\.c=\"" + col + "\"][z\\\\.r=\"" + row + "\"]").attr("style")
    	+jq("div.zscell[z\\\\.c=\"" + col + "\"][z\\\\.r=\"" + row + "\"] div").attr("style");
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
    	result += jq("div.zscell[z\\\\.c=\"" + col + "\"][z\\\\.r=\"" + row + "\"]").css("background-color");
    	result += ":"+jq("div.zscell[z\\\\.c=\"" + col + "\"][z\\\\.r=\"" + row + "\"] div").css("font-family");
    	result += ":"+jq("div.zscell[z\\\\.c=\"" + col + "\"][z\\\\.r=\"" + row + "\"] div").css("text-align");
    	return result;
    }
    /**
     * 
     * @param col - Base on 0
     * @param row - Base on 0
     * @return A JQuery object of cell. (Inner div)
     */
    public JQuery getSpecifiedCell(int col, int row) {
    	//for not IE
        return jq("div.zscell[z\\\\.c=\"" + col + "\"][z\\\\.r=\"" + row + "\"] div").first();
    }

    /**
     * 
     * @param col - Base on 0
     * @param row - Base on 0
     * @return A JQuery object of cell Outer div.
     */    
    public JQuery getSpecifiedCellOuter(int col, int row) {
        return jq("div.zscell[z\\\\.c=\"" + col + "\"][z\\\\.r=\"" + row + "\"]");
    }

    /**
     * 
     * @param col - Base on 0
     * @return A JQuery object of column header.
     */
    public JQuery getColumnHeader(int col) {
        return jq("div.zstopcell[z\\\\.c=\"" + col + "\"] div");
    }
    
    /**
     * 
     * @param row - Base on 0
     * @return A JQuery object of row header.
     */
    public JQuery getRowHeader(int row) {
       return jq("div.zsleftcell[z\\\\.r=\"" + row + "\"] div"); 
    }
    
    public String getCellContent(JQuery cellLocator) {
        return cellLocator.text();
    }
    
    public void clickCell(JQuery cellLocator) {
        mouseDownAt(cellLocator, "1,2");
        waitResponse();
        mouseUpAt(cellLocator, "1,2");
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
		mouseDownAt(getSpecifiedCell(column, row),"2,2");
		waitResponse();
		mouseUpAt(getSpecifiedCell(column, row),"2,2");
		waitResponse();
	
		mouseDownAt(getSpecifiedCell(column, row),"2,2");
		waitResponse();
		mouseUpAt(getSpecifiedCell(column, row),"2,2");
		waitResponse();

		contextMenuAt(getSpecifiedCell(column, row), "2,2");
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
		mouseDownAt(getSpecifiedCell(left, top),"2,2");
		waitResponse();
		mouseUpAt(getSpecifiedCell(left, top),"2,2");
		waitResponse();
	
		mouseDownAt(getSpecifiedCell(left, top),"2,2");
		waitResponse();
		mouseMoveAt(getSpecifiedCell(right, bottom),"2,2");
		waitResponse();
		mouseUpAt(getSpecifiedCell(right, bottom),"2,2");
		waitResponse();
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
}
