import org.zkoss.ztl.JQuery;
import org.zkoss.ztl.util.ColorVerifingHelper;

//Format>> Background color

public class SS_033_Test extends SSAbstractTestCase {

    @Override
    protected void executeTest() {
        JQuery cell_B_8 = getSpecifiedCell(1, 7);
        clickCell(cell_B_8);
        clickCell(cell_B_8);

    	click("jq('$formatMenu button.z-menu-btn')");
    	waitResponse();
    	click(jq("$backgroundColorMenu a.z-menu-cnt-img"));
    	waitResponse();

    	// Input color hex code, then press Enter.
    	//JQuery colorTextbox = jq(".z-colorpalette-hex-inp:visible");
    	//String backgroundColorStr = "#990033";
        //type(colorTextbox, backgroundColorStr);
        //keyPressEnter(colorTextbox); //selenium key press enter does not work well
        
    	JQuery color = jq("$backgroundColorMenu .z-colorpalette div.z-colorpalette-colorbox:nth-child(98)");
    	mouseOver(color);
    	click(color);
    	waitResponse();
    	String selectedColor = color.first().text();

        //Verify
        cell_B_8 = getSpecifiedCell(1, 7);
        String bgB8 = cell_B_8.parent().css("background-color");
        String style = ColorVerifingHelper.transform(bgB8);
        
        if (style != null) {
        	verifyTrue("Cell background color shall be: " + selectedColor, ColorVerifingHelper.isEqualColor(selectedColor, style));
        } else {
            verifyTrue("Cannot get style of specified cell!", false);
        }
   }
}
