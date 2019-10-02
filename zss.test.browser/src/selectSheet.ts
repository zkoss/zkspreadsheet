import Spreadsheet from './pageObject/Spreadsheet';

fixture `select sheets`
    .page `http://localhost:8080/zss.test/issue.zul?file=/book/t10425.xlsx`;

    
/**
 * select sheets one by one for nTimes
 */
test('select sheets', async t => {
   let spreadsheet = new Spreadsheet();

    var nTimes = 10;
    for (var c = 0 ; c < nTimes ; c++){
        for (var i=0 ; i < 6 ; i++){
            await spreadsheet.select(i);
        }
    }
});

