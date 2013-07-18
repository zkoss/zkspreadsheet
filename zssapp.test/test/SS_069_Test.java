import org.zkoss.ztl.util.ColorVerifingHelper;


public class SS_069_Test extends SSAbstractTestCase {

	/**
	 * Set bottom border
	 */
    @Override
    protected void executeTest() {
    	focusOnCell(11, 12);
    	clickDropdownButtonMenu("$fastIconBtn $borderBtn", "Bottom border");
    	String color = getSpecifiedCellOuter(11, 12).css("border-bottom-color");
        verifyTrue(ColorVerifingHelper.isEqualColor("#000000", color));
    }

}
