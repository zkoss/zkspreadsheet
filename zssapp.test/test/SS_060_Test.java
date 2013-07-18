import org.zkoss.ztl.JQuery;


public class SS_060_Test extends SSAbstractTestCase {

	/**
	 * Sets font family
	 */
    @Override
    protected void executeTest() {
    	focusOnCell(1, 7);
        click(jq("$fontCtrlPanel $fontFamily:visible i.z-combobox-rounded-btn"));
        JQuery fontItem = jq(".z-combobox-rounded-pp .timeFont:visible");
        
        click(fontItem);
        String cellFont = getCellFontFamily(1, 7);
        verifyTrue("Unexcepted result: " + cellFont, "Times New Roman".equalsIgnoreCase(cellFont));
    }

}
