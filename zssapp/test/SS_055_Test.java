import java.util.Map;

import org.zkoss.ztl.JQuery;


public class SS_055_Test extends SSAbstractTestCase {

    @Override
    protected void executeTest() {
        // Select source cell
        JQuery cell_J_13 = getSpecifiedCell(9, 12);
        Map<String, String> sourceStyleMap = getCellStyleMap(cell_J_13);
        String sourceBorderBottomStyle = sourceStyleMap.get("border-bottom");
        String sourceBorderRightStyle = sourceStyleMap.get("border-right");
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
        JQuery cell_L_13 = loadTargetCell();
        clickCell(cell_L_13);

        // Click Paste icon
        mouseOver(jq("$pasteDropdownBtn"));
        clickAt(jq("$pasteDropdownBtn"), "30,2");
        waitResponse();
        
        // Click Paste all except border
        click(jq("$pasteAllExcpetBorder"));
        waitResponse();
        
        // Verify
        cell_L_13 = loadTargetCell();
        Map<String, String> targetStyleMap = getCellStyleMap(cell_L_13);
        verifyNotEquals("border-bottom are not the same.", sourceBorderBottomStyle, targetStyleMap.get("border-bottom"));
        verifyNotEquals("border-right are not the same.", sourceBorderRightStyle, targetStyleMap.get("border-right"));
    }
    
    private JQuery loadTargetCell() {
        return getSpecifiedCell(11, 12);
    }

}
