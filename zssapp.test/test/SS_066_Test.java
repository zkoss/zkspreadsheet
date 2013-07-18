import org.zkoss.ztl.util.ColorVerifingHelper;


public class SS_066_Test extends SSAbstractTestCase {

	/**
	 * Sets cell font color from toolbar
	 */
    @Override
    protected void executeTest() {
        focusOnCell(1, 7);
        
        // Click font color button on the toolbar
        String selectedColor = setCellFontColorByToolbarbutton(1, 7, 98);
        //Verify
        verifyTrue(ColorVerifingHelper.isEqualColor(selectedColor, getCellFontColor(1, 7)));
    }

}
