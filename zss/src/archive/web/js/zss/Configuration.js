/* Configuration.js

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Mon Apr 23, 2007 17:29:18 AM , Created by sam
}}IS_NOTE

Copyright (C) 2007 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
	This program is distributed under GPL Version 2.0 in the hope that
	it will be useful, but WITHOUT ANY WARRANTY.
}}IS_RIGHT
*/

/**
 * A Configuration is used for spreadsheet configuration setting
 */
zss.Configuration = zk.$extends(zk.Object, {
	textOverflow: true,
	prune: true,
	readonly: false,
	hideBorder: false
});
zss.clientCopy = {
	maxRowCount: 1000,
	maxColumnCount: 40
};
zss.SCROLL_DIR = { // ZSS-475: add direction enum. for scrolling to visible
	BOTH: "both",
	HORIZONTAL: "horizontal",
	VERTICAL: "vertical"
};
zss.SCROLL_POS = { // ZSS-664: add position enum. for scrolling to position
	NONE: 'none',
	TOP: 'top',
	BOTTOM: 'bottom'
};