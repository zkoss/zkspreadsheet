import { Selector, t } from 'testcafe';

export class AutoFilter{
    spreadsheetSelector: any;
    filterButtons: any;
    colorFilter: any;

    constructor (spreadsheetSelector: any){
        this.spreadsheetSelector = spreadsheetSelector;
        this.filterButtons = spreadsheetSelector.find('.zsblock .zsdropdown');
        this.colorFilter = spreadsheetSelector.find('.zsafp-color');
    }

    /**
     * @param index 0-based filter index, from left to right
     */
    async openFilter(index: number){
        if (await this.filterButtons.exists){
            await t.click(this.filterButtons.nth(index));
        }
    }

    async isColorFilterEnabled(){
        let disabled = await this.spreadsheetSelector.find('.zsafp-color').hasClass('zsafp-color-disabled');
        return !disabled;
    }

    // expand font/background color items
    async selectColorFilter(){
        await t.click(this.spreadsheetSelector.find('.zsafp-color'));
    }  

    async getFontColorCount(){
        await this.selectColorFilter();
        return this.spreadsheetSelector.find('.zsafp-ccitem').count;
    }

    async getBackgroundColorCount(){
        await this.selectColorFilter();
        return this.spreadsheetSelector.find('.zsafp-fcitem').count; 
    }
    
    async selectNumberFilter(){
        await t.click(this.spreadsheetSelector.find('.zsafp-value'));
    }  
    async selectNumberFilterEquals(){
        await t.click(this.spreadsheetSelector.find('.zsafp-valuedlg .zsafp-vitem').nth(0));
    }  
}