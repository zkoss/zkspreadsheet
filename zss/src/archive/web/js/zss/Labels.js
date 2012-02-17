/* Label.js

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

	US = {
		_sheetSheet: 'Sheet',
		_sheetAdd: 'Add sheet',
		_sheetDelete: 'Delete',
		_sheetRename: 'Rename',
		_sheetMoveLeft: 'Move left',
		_sheetMoveRight: 'Move right',
		_sheetProtect: 'Protect sheet'
	};
	
zss.Labels = zk.$extends(zk.Object, {
	//TODO: Locale-dependent label
	$init: function (locale) {
		zk.copy(this, US);
	},
	$define: {
		sheetSheet: null,
		sheetAdd: null,
		sheetDelete: null,
		sheetRename: null,
		sheetMoveLeft: null,
		sheetMoveRight: null,
		sheetProtect: null
	}
});
})();