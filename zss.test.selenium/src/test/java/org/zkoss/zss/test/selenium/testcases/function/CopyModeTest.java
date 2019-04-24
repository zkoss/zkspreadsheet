package org.zkoss.zss.test.selenium.testcases.function;

import org.junit.*;
import org.openqa.selenium.Keys;
import org.zkoss.zss.test.selenium.*;
import org.zkoss.zss.test.selenium.entity.*;

/**
 * FIXME only test with chrome
 */
public class CopyModeTest extends ZSSTestCase {


    @Test
    public void copy1CellHighlight() {
        getTo("/blank.zul");
        sheetFunction().copy();
        waitForTime(Setup.getTimeoutL0());
        JQuery $highlight = jq(".zshighlight");
        Assert.assertEquals("block", $highlight.css("display"));
        Assert.assertEquals("35px", $highlight.css("left"));
        Assert.assertEquals("20px", $highlight.css("top"));
        Assert.assertEquals(59, $highlight.width());
        Assert.assertEquals(16, $highlight.height());
    }

    @Test
    public void copy4CellsHighlight() {
        getTo("/blank.zul");
        SpreadsheetWidget ss = focusSheet();
        SheetCtrlWidget ctrl = ss.getSheetCtrl();
        dragAndDrop(ctrl.getCell("A1").toWebElement(), ctrl.getCell("B2").toWebElement());
        sheetFunction().copy();
        waitForTime(Setup.getTimeoutL0());
        JQuery $highlight = jq(".zshighlight");
        Assert.assertEquals("block", $highlight.css("display"));
        Assert.assertEquals("35px", $highlight.css("left"));
        Assert.assertEquals("20px", $highlight.css("top"));
        Assert.assertEquals(123, $highlight.width());
        Assert.assertEquals(37, $highlight.height());
    }


    /**
     * cancel copy mode after editing
     */
    @Test
    public void editCancelCopyMode() {
        getTo("/blank.zul");
        sheetFunction().copy();
        waitForTime(Setup.getTimeoutL0());
        JQuery $highlight = jq(".zshighlight");
        Assert.assertEquals("block", $highlight.css("display"));
        Assert.assertEquals("35px", $highlight.css("left"));
        Assert.assertEquals("20px", $highlight.css("top"));
        Assert.assertEquals(59, $highlight.width());
        Assert.assertEquals(16, $highlight.height());

        SpreadsheetWidget ss = focusSheet();
        SheetCtrlWidget sheetCtrl = ss.getSheetCtrl();
        click(sheetCtrl.getCell("A2"));
        EditorWidget editor = sheetCtrl.getInlineEditor();
        sendKeys(editor, "any");
        waitForTime(Setup.getTimeoutL0());
        Assert.assertEquals("none", $highlight.css("display"));
    }

    /**
     * pasting range doesn't overlap the source range, the copy mode keeps
     */
    @Test
    public void pasteKeepsCopyMode() {
        getTo("/blank.zul");
        SpreadsheetWidget ss = focusSheet();
        SheetCtrlWidget sheetCtrl = ss.getSheetCtrl();
        EditorWidget editor = sheetCtrl.getInlineEditor();
        doubleClick(sheetCtrl.getCell("A1"));
        sendKeys(editor, "a");
        sendKeys(editor, Keys.chord(Keys.ENTER));
        doubleClick(sheetCtrl.getCell("A2"));
        sendKeys(editor, "b");
        sendKeys(editor, Keys.chord(Keys.ENTER));
        doubleClick(sheetCtrl.getCell("B1"));
        sendKeys(editor, "c");
        sendKeys(editor, Keys.chord(Keys.ENTER));
        doubleClick(sheetCtrl.getCell("B2"));
        sendKeys(editor, "d");
        sendKeys(editor, Keys.chord(Keys.ENTER));
        dragAndDrop(sheetCtrl.getCell("A1").toWebElement(), sheetCtrl.getCell("B2").toWebElement());
        sheetFunction().copy();
        click(sheetCtrl.getCell("C1"));
        sheetFunction().paste();
        JQuery $highlight = jq(".zshighlight");
        Assert.assertEquals("block", $highlight.css("display"));
        Assert.assertEquals("35px", $highlight.css("left"));
        Assert.assertEquals("20px", $highlight.css("top"));
        Assert.assertEquals(123, $highlight.width());
        Assert.assertEquals(37, $highlight.height());
    }

    /**
     * pasting range overlaps the source range, the copy mode cancels
     */
    @Test
    public void pasteCancelsCopyMode() {
        getTo("/blank.zul");
        SpreadsheetWidget ss = focusSheet();
        SheetCtrlWidget sheetCtrl = ss.getSheetCtrl();
        EditorWidget editor = sheetCtrl.getInlineEditor();
        doubleClick(sheetCtrl.getCell("A1"));
        sendKeys(editor, "a");
        sendKeys(editor, Keys.chord(Keys.ENTER));
        doubleClick(sheetCtrl.getCell("A2"));
        sendKeys(editor, "b");
        sendKeys(editor, Keys.chord(Keys.ENTER));
        doubleClick(sheetCtrl.getCell("B1"));
        sendKeys(editor, "c");
        sendKeys(editor, Keys.chord(Keys.ENTER));
        doubleClick(sheetCtrl.getCell("B2"));
        sendKeys(editor, "d");
        sendKeys(editor, Keys.chord(Keys.ENTER));
        dragAndDrop(sheetCtrl.getCell("A1").toWebElement(), sheetCtrl.getCell("B2").toWebElement());
        sheetFunction().copy();
        click(sheetCtrl.getCell("B1"));
        sheetFunction().paste();
        waitForTime(Setup.getTimeoutL0());
        JQuery $highlight = jq(".zshighlight");
        Assert.assertEquals("none", $highlight.css("display"));
    }

}





