import { Selector } from 'testcafe';
import Spreadsheet from './pageObject/Spreadsheet';

//ZSS-1401 a number filter cause null pointer exception
fixture`color filter`
    .page`http://localhost:8080/zss.test/issue.zul?file=/book/table-filter.xlsx`;

test('number filter available', async t => {
    var spreadsheet = new Spreadsheet();
    const autoFilterDialog = Selector('.z-window-modal');

    await spreadsheet.autoFilter.openFilter(0);
    await spreadsheet.autoFilter.selectNumberFilter();
    await spreadsheet.autoFilter.selectNumberFilterEquals();
    await t.expect(autoFilterDialog.exists).ok();
});


