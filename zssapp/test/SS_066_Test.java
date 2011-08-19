import org.zkoss.ztl.JQuery;
import org.zkoss.ztl.util.ColorVerifingHelper;


public class SS_066_Test extends SSAbstractTestCase {

	/**
	 * Sets cell font color from toolbar
	 */
    @Override
    protected void executeTest() {
        focusOnCell(1, 7);
        
        // Click font color button on the toolbar
        click(jq("$fontCtrlPanel $fontColorBtn"));
        waitResponse();
        
        JQuery color = jq(".z-colorpalette:visible div.z-colorpalette-colorbox:nth-child(98)");
        String selectedColor = color.first().text();
    	mouseOver(color);
    	click(color);
    	waitResponse();
        
        //Verify
        verifyTrue(ColorVerifingHelper.isEqualColor(selectedColor, getCellFontColor(1, 7)));
    }

}
