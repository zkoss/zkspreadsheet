import { Selector, t } from 'testcafe';
import { AutoFilter } from './AutoFilter';

/**
 * Implement a page object pattern.
 * Reference:
 * - https://devexpress.github.io/testcafe/documentation/recipes/extract-reusable-test-code/use-page-model.html
 * - https://martinfowler.com/bliki/PageObject.html
 */
export default class Spreadsheet{
    selector: any;
    sheetbar: any;
    autoFilter: AutoFilter;

    constructor (){
        this.selector = Selector('.zssheet');
        this.sheetbar = this.selector.find('.zssheettab');
        this.autoFilter = new AutoFilter(this.selector);
        this.focus();
    }

    async select(index: number){
        await t.click(this.sheetbar.nth(index));
    }
    
    async focus(){
        await t.click(this.selector);
    }
}
