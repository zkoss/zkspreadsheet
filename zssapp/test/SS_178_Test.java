import org.zkoss.ztl.JQuery;


public class SS_178_Test extends SSAbstractTestCase {

    @Override
    protected void executeTest() {
        // Select source cell
    	focusOnCell(9, 12);
    	String sourceBorderBottomStyle = getCellBorderStyle(9, 12, BORDER_BOTTOM);
    	String sourceBorderRightStyle = getCellBorderStyle(9, 12, BORDER_RIGHT);
        
        // Ctrl + C
        keyDownNative(CTRL);
        waitResponse();
        keyDownNative(C);
        waitResponse();
        keyUpNative(C);
        waitResponse();
        keyUpNative(CTRL);
        waitResponse();

        JQuery L13 = rightClickCell(11, 12);
        waitResponse();
        
        // Click Paste Special on the context menu
        click(jq("$pasteSpecial"));
        waitResponse();
        
        // Choose All except border
        click(jq("$allExcpetBorder input[id*=real]"));
        waitResponse();
        click(jq("$okBtn"));
        waitResponse();
        
        // Verify
        verifyNotEquals("border-bottom are not the same.", sourceBorderBottomStyle, getCellBorderStyle(L13, BORDER_BOTTOM));
        verifyNotEquals("border-right are not the same.", sourceBorderRightStyle, getCellBorderStyle(L13, BORDER_RIGHT));
    }
}
