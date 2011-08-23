import java.util.Random;

import org.zkoss.ztl.JQuery;


public class SS_002_Test extends SSAbstractTestCase {

	/**
	 * Open a empty spreadsheet
	 */
    @Override
    protected void executeTest() {
    	click("$fileMenu");
    	waitResponse();
    	click("$newFile:visible");
        /**
         * Expect:
         * 
         * spreadsheet mask is invisible, cell text shall be empty
         */
        verifyFalse("mask shall be invisible", isWidgetVisible(".zssmask")) ;

    	Random randomGenerator = new Random();
        JQuery newc = getSpecifiedCell(randomGenerator.nextInt(25), randomGenerator.nextInt(25));
        String newcValue = getCellText(newc);
        verifyTrue("".equals(newcValue) || newc == null);
    }
}
