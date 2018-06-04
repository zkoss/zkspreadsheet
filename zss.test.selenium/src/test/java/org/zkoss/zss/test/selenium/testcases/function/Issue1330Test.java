package org.zkoss.zss.test.selenium.testcases.function;

import org.junit.*;
import org.openqa.selenium.*;
import org.zkoss.zss.test.selenium.*;
import org.zkoss.zss.test.selenium.entity.*;

import static org.junit.Assert.*;


public class Issue1330Test extends ZSSTestCase {

    /*
     * verify the importing result
     */
    @Test
    public void testZSS1338Import() throws Exception {
        getTo("/issue3/1338-overflow-right-align.zul");
        basename();
        captureOrAssert("import");
    }

    // always move selection box to A1 before capturing a screenshot to avoid image comparison being affected by selection box
    @Test
    public void testZSS1338ChangeAlign() throws Exception {
        getTo("/issue3/1338-overflow-right-align.zul");
        basename();
        SpreadsheetWidget ss = focusSheet();
        SheetCtrlWidget ctrl = ss.getSheetCtrl();
        click(ctrl.getCell("F5"));
        dragAndDrop(ctrl.getCell("F5").toWebElement(), ctrl.getCell("F10").toWebElement());
        click(ZSStyle.HORIZONTAL_ALIGN + " .zstbtn-cave");
        click(ZSStyle.MENUPOPUP_OPEN.toString() + " " + ZSStyle.ALIGN_LEFT.toString());
        click(ctrl.getCell("A1"));
        captureOrAssert("changeToLeftAlign");

        dragAndDrop(ctrl.getCell("F5").toWebElement(), ctrl.getCell("F10").toWebElement());
        click(ZSStyle.HORIZONTAL_ALIGN + " .zstbtn-cave");
        click(ZSStyle.MENUPOPUP_OPEN.toString() + " " + ZSStyle.ALIGN_RIGHT.toString());
        click(ctrl.getCell("A1"));
        captureOrAssert("changeToRightAlign");
    }

    @Test
    public void testZSS1338PartialOverflow() throws Exception {
        getTo("/issue3/1338-overflow-right-align.zul");
        basename();
        SpreadsheetWidget ss = focusSheet();
        SheetCtrlWidget sheetCtrl = ss.getSheetCtrl();
        click(sheetCtrl.getCell("D9"));
        doubleClick(sheetCtrl.getCell("D9"));
        waitForTime(Setup.getTimeoutL0());
        EditorWidget editor = sheetCtrl.getInlineEditor();
        sendKeys(editor, "cut");
        click(sheetCtrl.getCell("A1"));
        captureOrAssert("partial1");
        click(sheetCtrl.getCell("E9"));
        doubleClick(sheetCtrl.getCell("E9"));
        sendKeys(editor, "cut");
        click(sheetCtrl.getCell("A1"));
        captureOrAssert("partial2");
    }

    @Test
    public void testZSS1364SheetSwitching(){
        getTo("/issue3/1338-overflow-right-align.zul");
        basename();
        sheetFunction().gotoTab(2);
        waitForTime(Setup.getTimeoutL0());
        sheetFunction().gotoTab(1);
        captureOrAssert("1");
    }
}





