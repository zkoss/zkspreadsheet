import { Selector } from 'testcafe';
import { ClientFunction } from 'testcafe';

fixture `select sheets`
    .page `http://localhost:8080/zss.test/issue.zul?file=/book/t10425.xlsx`;

    
const tabs = Selector('.zssheettab'); 

/**
 * select sheets one by one for nTimes
 */
test('select sheets', async t => {
   
    var nTimes = 10;
    for (var c = 0 ; c < nTimes ; c++){
        for (var i=0 ; i < 6 ; i++){
            await t.click(tabs.nth(i));
        }
    }
});

