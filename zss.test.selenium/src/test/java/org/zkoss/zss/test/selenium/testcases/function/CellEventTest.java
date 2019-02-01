package org.zkoss.zss.test.selenium.testcases.function;

import org.junit.*;
import org.openqa.selenium.Keys;
import org.zkoss.zss.test.selenium.*;
import org.zkoss.zss.test.selenium.entity.*;
import org.zkoss.zss.test.selenium.util.MyPngCropper;

import static org.zkoss.zss.test.selenium.util.MyPngCropper.TOP_BOTTOM_CROPPER;


public class CellEventTest extends ZSSTestCase {

    public static final String ON_CELL_CLICK = "onCellClick click ";

    /**
     * click a merged cell, fires an event to the server.
     */
    @Test
    public void testZSS1377ClickMergedCellColumnA() {
        final String A1 = "A1";

        getTo("/issue3/231-cellclick-merge.zul");
        SpreadsheetWidget ss = focusSheet();
        SheetCtrlWidget ctrl = ss.getSheetCtrl();
        click(ctrl.getCell(A1));
        waitForTime(Setup.getTimeoutL0());
        Assert.assertEquals(ON_CELL_CLICK + A1, jq("$msg").toWebElement().getText());
    }

    /**
     * click a merged cell, the server side gets the left-top corner.
     */
    @Test
    public void testZSS231ClickMerged() {
        getTo("/issue3/231-cellclick-merge.zul");
        SpreadsheetWidget ss = focusSheet();
        SheetCtrlWidget ctrl = ss.getSheetCtrl();
        click(ctrl.getCell("C3"));
        waitForTime(Setup.getTimeoutL0());
        Assert.assertEquals(ON_CELL_CLICK + "B2", jq("$msg").toWebElement().getText());

        click(ctrl.getCell("D8"));
        waitForTime(Setup.getTimeoutL0());
        Assert.assertEquals(ON_CELL_CLICK + "D5", jq("$msg").toWebElement().getText());
    }
}





