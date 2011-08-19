import org.zkoss.ztl.JQuery;
import org.zkoss.ztl.util.ColorVerifingHelper;


public class SS_067_Test extends SSAbstractTestCase {

    @Override
    protected void executeTest() {
        focusOnCell(1, 7);
        
        // Click font color button on the toolbar
        click(jq("$fontCtrlPanel $cellColorBtn"));
        waitResponse();
        
        JQuery color = jq(".z-colorpalette:visible div.z-colorpalette-colorbox:nth-child(98)");
        //TODO: decode #9966ff format to rgb format
        //String selectedColor = color.first().text();
    	mouseOver(color);
    	click(color);
    	
        //Verify
        verifyEquals("rgb(153, 102, 255)", getCellBackgroundColor(1, 7));
    }
}
