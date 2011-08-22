import org.zkoss.ztl.JQuery;

import com.google.common.base.Objects;


public class SS_054_Test extends SSAbstractTestCase {

	/**
	 * Paste (value only)
	 */
    @Override
    protected void executeTest() {
    	JQuery srcFormulaCell = focusOnCell(5, 14);
    	String formulaBarValue = jq("$formulaEditor").val();
    	String srcCellText = getCellText(srcFormulaCell);
    	verifyTrue("Shall select formula cell as paste source", 
    		formulaBarValue.startsWith("=") && !Objects.equal(srcFormulaCell, srcCellText));
    	verifyTrue("Shall select format cell as paste source", srcCellText.indexOf(",") >= 0);
    	click("$copyBtn");
    	
    	//paste value
    	focusOnCell(12, 14);
    	clickDropdownButtonMenu("$fastIconBtn $pasteDropdownBtn", "Values");
    	String targetFormulaBarValue = jq("$formulaEditor").val();
    	/**
    	 * Paste value, no style, no format
    	 */
    	verifyTrue("paste value: shall not be formula", !(targetFormulaBarValue.indexOf("=") >= 0));
    	srcCellText = srcCellText.replace(",", "");
    	verifyEquals("paste value, no format", srcCellText, targetFormulaBarValue);
    }
}
