import org.zkoss.ztl.JQuery;


public class SS_055_Test extends SSAbstractTestCase {

	/**
	 * Paste All expect border
	 */
    @Override
    protected void executeTest() {
        // Select source cell
        JQuery cell_J_13 = getSpecifiedCell(9, 12);
        String sourceBorderBottomStyle = getCellBorderStyle(9, 12, BORDER_BOTTOM);
        String sourceBorderRightStyle = getCellBorderStyle(9, 12, BORDER_RIGHT);
        clickCell(cell_J_13);
        clickCell(cell_J_13);
        
        // Ctrl + C
        keyDownNative(CTRL);
        waitResponse();
        keyDownNative(C);
        waitResponse();
        keyUpNative(C);
        waitResponse();
        keyUpNative(CTRL);
        waitResponse();
        
        // Right click target cell
        focusOnCell(11, 12);
        
        // Click Paste icon
        mouseOver(jq("$pasteDropdownBtn"));
        clickAt(jq("$pasteDropdownBtn"), "30,2");
        waitResponse();
        
        // Click Paste all except border
        click(jq("$pasteAllExcpetBorder"));
        waitResponse();
        
        // Verify
        String targetBorderBottom = getCellBorderStyle(11, 12, BORDER_BOTTOM);
        String targetBorderRIGHT = getCellBorderStyle(11, 12, BORDER_RIGHT);
        verifyNotEquals("border-bottom are not the same.", sourceBorderBottomStyle, targetBorderBottom);
        verifyNotEquals("border-right are not the same.", sourceBorderRightStyle, targetBorderRIGHT);
    }
}
