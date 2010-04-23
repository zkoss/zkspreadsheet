package org.zkoss.zss.demo;
import org.zkoss.zk.ui.*;
import org.zkoss.zk.ui.ext.AfterCompose;
import org.zkoss.zul.*;

/**
 * The Table-of-Content tree on the left.
 *
 * 
 */
public class zssTocTree extends Tree implements AfterCompose {
	public zssTocTree() {
	}
	public void onSelect() {
		Treeitem item = getSelectedItem();
		if (item != null) {
			Include inc = (Include)Path.getComponent("/showRoom/contentArea");
			System.out.println("src: " + item.getValue());
			inc.setSrc((String)item.getValue());
		}
	}
	public void afterCompose() {
		final Execution exec = Executions.getCurrent();
		String id = exec.getParameter("id");
		Treeitem item = null;
		if (id != null) {
			try {
				item = (Treeitem)getSpaceOwner().getFellow(id);
			} catch (ComponentNotFoundException ex) { //ignore
			}
		}

		if (item == null) {
			item = (Treeitem)getSpaceOwner().getFellow("ui1");
			exec.setAttribute("contentSrc", (String)item.getValue());
			selectItem(item);
		}
	}
}
