/* CacheCtrl.js

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Jan 12, 2012 7:07:27 PM , Created by sam
}}IS_NOTE

Copyright (C) 2012 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
(function () {
	


/**
 * Control client cache data of sheets
 */
zss.CacheCtrl = zk.$extends(zk.Object, {
	/**
	 * Current sheet data
	 */
	selected: null,
	$init: function (spreadsheet, v) {
		this._wgt = spreadsheet;
		this.sheet = spreadsheet.sheetCtrl;
		
		//key: sheet uuid, value: zss.ActiveRange
		this.sheets = {};
		
		//zss.Snapshot. key: sheet uuid, value: sheet last status
		this.snapshots = {};
		
		this.setSelectedSheet(v);
	},
	/**
	 * Save current sheet status
	 */
	snap: function (sheetId) {
		return this.snapshots[sheetId] = new zss.Snapshot(this._wgt);
	},
	getSnapshot: function (sheetId) {
		return this.snapshots[sheetId];
	},
	isCached: function (sheetId) {
		return !!this.sheets[sheetId];
	},
	releaseCache: function (sheetId) {
		if(this.sheets[sheetId] && this.sheets[sheetId] != this.selected){
			delete this.sheets[sheetId];
		}
		if(this.snapshots[sheetId]){
			delete this.snapshots[sheetId];
		}
	},
	//ZSS-1181
	releaseSelectedCache: function (sheetId) {
		if (this.selected && this.selected.id == sheetId) {
			delete this.sheets[sheetId];
			delete this.selected;
		}
		if(this.snapshots[sheetId]){
			delete this.snapshots[sheetId];
		}
	},
	setSelectedSheetBy: function (sheetId) {
		this.selected = this.sheets[sheetId];
	},
	getSheetBy: function (shtId) {
		return this.sheets[shtId];
	},
	setSelectedSheet: function (v) {
		var sheetId = v.id,
			rng = this.sheets[sheetId] = new zss.ActiveRange(v);
		
		this.selected = rng;
	},
	/** selected sheet might be inconsistent between server and client for network latency
	*/
	getSelectedSheet: function () {
		return this.selected;
	}
});


})();
