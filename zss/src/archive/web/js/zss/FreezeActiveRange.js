/* FreezeActiveRange.js

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Jan 12, 2012 7:07:27 PM , Created by sam
		Sep 12, 2019, separated by Hawk Chen
}}IS_NOTE

Copyright (C) 2012 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
(function () {

//ZSS-392: update freeze panels' activeRange individually
zss.FreezeActiveRange = zk.$extends(zss.ActiveRange, {

	$init: function (data) {
		this.$supers(zss.FreezeActiveRange, '$init', [data]); 
	},
	
	// override
	update: function (v, dir) {
		// just update cells
		this.updateCells(v, dir);
	},
	
	// ZSS-404: freeze panels should also update row/column
	// override
	insertNewColumn: function (colIdx, size, headers) {
		this.insertNewColumn_(colIdx, size, headers); // just update row/column
	},
	removeColumns: function (col, size, headers) {
		this.removeColumns_(col, size, headers);
	},
	insertNewRow: function (rowIdx, size, headers) {
		this.insertNewRow_(rowIdx, size, headers);
	},
	removeRows: function (row, size, headers) {
		this.removeRows_(row, size, headers);
	}
	
});

})();
