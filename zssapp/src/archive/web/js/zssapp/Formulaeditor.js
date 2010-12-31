/* Formulaeditor.js

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Dec 30, 2010 11:25:02 AM , Created by henrichen
}}IS_NOTE

Copyright (C) 2010 Potix Corporation. All Rights Reserved.

*/
(function () {

zssapp.Formulaeditor = zk.$extends(zul.inp.Textbox, {
	_newLine: function () {
		var editorcmp = this.$n(),
			sel = zk(editorcmp).getSelectionRange();	
		if (sel[0] > sel[1]) {
			var t = sel[1];
			sel[1] = sel[0];
			sel[0] = t;
		}
		var str = editorcmp.value,
			len = str.length,
			str1 = str.substring(0, sel[0]); 
		str = str1 +'\n'+ str.substring(sel[1], len);
		editorcmp.value = str;

		var pos = sel[0];
		//move text cursor position to last
		fun = function () {
			if (editorcmp.setSelectionRange) { 
				pos = zk.opera ? pos + 2 : pos + 1;
				editorcmp.setSelectionRange(pos, pos);
			} else if (editorcmp.createTextRange) {//IE
				var range = editorcmp.createTextRange(),
					i = -1;
				pos = pos + 1;
				while((i = str1.indexOf('\n',i+1)) >= 0) {
					pos--;
				}
				range.move('character', pos);
				range.select();
			}
		};
		fun();
	},
	doKeyDown_: function (evt) {
		var keyCode = evt.keyCode;
		if(keyCode == 13) { //
			if(!evt.altKey && !evt.ctrlKey) {
				if (this.beforeCtrlKeys_)
					this.beforeCtrlKeys_(evt);
				this.fire("onOK", zk.copy({reference: evt.target}, evt.data), {ctl: true});
			} else
				this._newLine();
			evt.stop();
		} else {
			if(keyCode == 9) //Tab
				this.fire("onTab", zk.copy({reference: evt.target}, evt.data), {ctl: true});
			this.$supers('doKeyDown_', arguments);
		}
	}
});
})();