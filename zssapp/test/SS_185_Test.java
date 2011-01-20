import java.util.Map;

import org.zkoss.ztl.JQuery;

public class SS_185_Test extends SSAbstractTestCase {

    @Override
    protected void executeTest() {
        // Select source cell
        JQuery cell_J_13 = getSpecifiedCell(9, 12);
        clickCell(cell_J_13);
        clickCell(cell_J_13);
        System.out.println(getCellStyleMap(getSpecifiedCell(14, 5)));

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
        rightClickCell(cell_L_13);
        waitResponse();

        // Click Paste Special on the context menu
        click(jq("$pasteSpecial"));
        waitResponse();

        // Choose Values and number formats
        click(jq("$fmt input[id*=real]"));
        waitResponse();
        click(jq("$okBtn"));
        waitResponse();

        // Verify
        Map<String, String> sourceStyleMap = getCellStyleMap(cell_J_13);
        Map<String, String> targetStyleMap = getCellStyleMap(cell_L_13);

        for (String key : sourceStyleMap.keySet()) {
        	if (!"border-right".equals(key)) {
        		verifyEquals(sourceStyleMap.get(key), targetStyleMap.get(key));
        	}
        }
    }

    private JQuery loadTargetCell() {
        return getSpecifiedCell(11, 12);
    }

}
