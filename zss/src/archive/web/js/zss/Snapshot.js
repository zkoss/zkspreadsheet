/* Snapshot.js

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
	    Jan 12, 2012 7:07:27 PM , Created by sam
		Sep 5, 2019, separated by Hawk chen
}}IS_NOTE

Copyright (C) 2012 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
(function () {
/**
 * Snapshot sheet UI status e.g. focus, scrolling position and attributes
 * 
 * <ul>
 * 	<li>sheet style</li>
 * 	<li>row freeze</li>
 * 	<li>column freeze</li>
 * 	<li>row height</li>
 * 	<li>column width</li>
 * 	<li>visible range</li>
 *  <li>focus</li>
 * 	<li>selection</li>
 *  <li>highlight</li>
 * 	<li>displayGridlines</li>
 * 	<li>protect</li>
 * </ul>
 */
zss.Snapshot = zk.$extends(zk.Object, {
	$init: function (spreadsheet) {
		var sheet = spreadsheet.sheetCtrl,
			dataPanel = sheet.dp,
			leftPanel = sheet.lp,
			topPanel = sheet.tp,
			scrollPanel = sheet.sp,
			visRng = zss.SSheetCtrl._getVisibleRange(sheet);
		zss.Snapshot.copyAttributes(this, spreadsheet,
			['_scss', '_displayGridlines', '_rowFreeze', '_columnFreeze', '_rowHeight', '_columnWidth', '_protect', '_maxRows', '_maxColumns']);
		
		this.setCustRowHeight(sheet.custRowHeight.custom);
		this.setCustRowLastId(sheet.custRowHeight.ids.last);
		this.setCustColWidth(sheet.custColWidth.custom);
		this.setCustColLastId(sheet.custColWidth.ids.last);
		this.setMergeMatrix(sheet.mergeMatrix.mergeMatrix);
		this.setVisibleRange(visRng);
		this.setFocus(sheet.getLastFocus());
		this.setSelection(sheet.getLastSelection());
		if (sheet.isHighlightVisible()) {
			this.setHighlight(sheet.getLastHighlight());
		}
		
		this.setDataPanelSize({'width': dataPanel.width, 'height': dataPanel.height});
		this.setScrollPanelPos({'scrollLeft': scrollPanel.currentLeft, 'scrollTop': scrollPanel.currentTop});
		this.setLeftPanelPos(leftPanel.toppos);
		this.setTopPanelPos(topPanel.leftpos);
		
		if (spreadsheet.getDataValidations) {
			var dv = spreadsheet.getDataValidations();
			if (dv) {
				this.setDataValidations(dv);
			}
		}
		if (spreadsheet.getAutoFilter) {
			var af = spreadsheet.getAutoFilter();
			if (af) {
				this.setAutoFilter(af);
			}
		}
		//ZSS-988
		if (spreadsheet.getTableFilters) {
			var tbafs = spreadsheet.getTableFilters();
			if (tbafs) {
				this.setTableFilters(tbafs);
			}
		}
	},
	$define: {
		scss: null,
		rowFreeze: null,
		columnFreeze: null,
		rowHeight: null,
		columnWidth: null,
		protect: null,
		displayGridlines: null,
		/**
		 * @param array
		 */
		custRowHeight: null,
		/**
		 * @param int
		 */
		custRowLastId:null,
		/**
		 * @param array
		 */
		custColWidth: null,
		/**
		 * @param int
		 */
		custColLastId:null,
		/**
		 * @param array
		 */
		mergeMatrix: null,
		visibleRange: null,
		/**
		 * Data panel's width/height
		 */
		dataPanelSize: null,
		/**
		 * Scroll panel's scroll left and scroll top position
		 */
		scrollPanelPos: null,
		/**
		 * Left panel's top position
		 */
		leftPanelPos: null,
		/**
		 * Top panel's left position
		 */
		topPanelPos: null,
		focus: null,
		selection: null,
		highlight: null,
		dataValidations: null,
		autoFilter: null,
		tableFilters: null, //ZSS-988
		maxRows: null, 		//ZSS-1082
		maxColumns: null 	//ZSS-1082
	}
},{ //static
    copyAttributes: function(dst, src, fields) {
        for (var key in fields) {
            var f = fields[key];
            dst[f] = src[f];
        }
    }
});
})();
