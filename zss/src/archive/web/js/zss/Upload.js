/* Upload.js

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
 * Upload 
 */
zss.Upload = zk.$extends(zk.Widget, {
	$init: function () {
		this.$supers(zss.Upload, '$init', [{visible: false}]);
	},
	getUploader: function (id) {
		var chd = this.firstChild;
		for (; chd; chd = chd.nextSibling) {
			if (chd.id == id) {
				return chd;
			}
		}
	}
});

zss.Uploader = zk.$extends(zk.Widget, {
	_upload: 'true',
	parasitize: function (wgt) {
		this.$n = function () {
			return wgt.$n();
		}
	}
});

})();