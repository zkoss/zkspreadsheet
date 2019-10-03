import { Selector } from 'testcafe';
import Spreadsheet from './pageObject/Spreadsheet';

//ZSS-1400 color filter causes null pointer exception
fixture`color filter`
    .page`http://localhost:8080/zss.test/issue.zul?file=/book/colorFilter.xlsx`;

test('disabled color filter', async t => {
    let spreadsheet = new Spreadsheet();
    await spreadsheet.autoFilter.openFilter(0);
    await t.expect(await spreadsheet.autoFilter.isColorFilterEnabled()).eql(false);
});

test('font color filter', async t => {
    let spreadsheet = new Spreadsheet();
    await spreadsheet.autoFilter.openFilter(2);
    await t.expect(await spreadsheet.autoFilter.isColorFilterEnabled()).eql(true);
    await t.expect(await spreadsheet.autoFilter.getFontColorCount()).eql(2);
});


test('background color filter', async t => {
    let spreadsheet = new Spreadsheet();
    await spreadsheet.autoFilter.openFilter(3);
    await t.expect(await spreadsheet.autoFilter.isColorFilterEnabled()).eql(true);
    await t.expect(await spreadsheet.autoFilter.getBackgroundColorCount()).eql(2);
});

