

public class SS_067_Test extends SSAbstractTestCase {

    @Override
    protected void executeTest() {
        focusOnCell(1, 7);
        
        // Click font color button on the toolbar
        String color = setCellBackgroundColorByToolbarbutton(1, 7, 98);
    	
        //TODO: compare hex format and rgb format
        //Verify
    	String bg = getCellBackgroundColor(1, 7);
        verifyTrue("rgb(153, 102, 255)".equals(bg) || "#9966cc".equals(bg));
    }
}
