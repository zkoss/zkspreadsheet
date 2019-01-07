package org.zkoss.zss.test.selenium.testcases.function;

import org.junit.*;
import org.openqa.selenium.*;
import org.zkoss.zss.test.selenium.*;
import org.zkoss.zss.test.selenium.entity.*;
import org.zkoss.zss.test.selenium.util.*;

import static org.zkoss.zss.test.selenium.util.MyPngCropper.TOP_BOTTOM_CROPPER;


public class Issue1300RightAlignOverflowTest extends ZSSTestCase {

    /*
     * verify the importing result
     */
    @Test
    public void testZSS1338Import() throws Exception {
        getTo("/issue3/1338-overflow-right-align.zul");
        basename();
        captureOrAssert("import", TOP_BOTTOM_CROPPER);
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
        focusSheet();
        captureOrAssert("changeToLeftAlign", TOP_BOTTOM_CROPPER);

        dragAndDrop(ctrl.getCell("F5").toWebElement(), ctrl.getCell("F10").toWebElement());
        click(ZSStyle.HORIZONTAL_ALIGN + " .zstbtn-cave");
        click(ZSStyle.MENUPOPUP_OPEN.toString() + " " + ZSStyle.ALIGN_RIGHT.toString());
        focusSheet();
        captureOrAssert("changeToRightAlign", TOP_BOTTOM_CROPPER);
    }

    @Test
    public void testZSS1338ChangeAlignMergedCell() throws Exception {
        getTo("/issue3/1338-overflow-right-align.zul");
        basename();
        SpreadsheetWidget ss = focusSheet();
        SheetCtrlWidget ctrl = ss.getSheetCtrl();
        click(ctrl.getCell("E11"));
        click(ZSStyle.HORIZONTAL_ALIGN + " .zstbtn-cave");
        click(ZSStyle.MENUPOPUP_OPEN.toString() + " " + ZSStyle.ALIGN_LEFT.toString());
        focusSheet();
        captureOrAssert("changeToLeftAlign", TOP_BOTTOM_CROPPER);

        click(ZSStyle.HORIZONTAL_ALIGN + " .zstbtn-cave");
        click(ZSStyle.MENUPOPUP_OPEN.toString() + " " + ZSStyle.ALIGN_RIGHT.toString());
        focusSheet();
        captureOrAssert("changeToRightAlign", TOP_BOTTOM_CROPPER);
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
        waitForTime(Setup.getTimeoutL0());
        captureOrAssert("1", TOP_BOTTOM_CROPPER);

    }

    @Test
    public void testZSS1364EditSheetSwitching() {
        getTo("/issue3/1338-overflow-right-align.zul");
        basename();
        SpreadsheetWidget ss = focusSheet();
        SheetCtrlWidget sheetCtrl = ss.getSheetCtrl();
        doubleClick(sheetCtrl.getCell("D5"));
        waitForTime(Setup.getTimeoutL0());
        EditorWidget editor = sheetCtrl.getInlineEditor();
        sendKeys(editor, "cut");
        sendKeys(editor, Keys.chord(Keys.ENTER));
        sheetFunction().gotoTab(2);
        waitForTime(Setup.getTimeoutL0());
        sheetFunction().gotoTab(1);
        captureOrAssert("1", TOP_BOTTOM_CROPPER);
    }

    @Test //ZSS-1364
    public void notRenderedCellAtFirst(){
        getTo("/issue3/1338-overflow-right-align.zul");
        basename();
        SpreadsheetWidget ss = focusSheet();
        setZSSScrollTop(1400);
        waitForTime(Setup.getTimeoutL1());
        captureOrAssert("1", new MyPngCropper(100, 600, 400, 0));
    }

    @Test //ZSS-1375
    public void resizeColumnProtectedSheet(){
        getTo("/issue3/1338-overflow-right-align.zul");
        basename();
        SpreadsheetWidget ss = focusSheet();
        sheetFunction().gotoTab(3);
        waitForTime(Setup.getTimeoutL1());
        //enlarge a column width
        dragAndDropBy(jq(".zshbouni").get(5), 100, 0);
        click(ss.getSheetCtrl().getCell("A1"));
        waitForTime(Setup.getTimeoutL1());
        captureOrAssert("1", TOP_BOTTOM_CROPPER);
    }

}





