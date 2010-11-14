package org.zkoss.zss.app.ctrl;

import org.zkoss.zk.ui.Component;
import org.zkoss.zk.ui.Page;
import org.zkoss.zk.ui.metainfo.ComponentInfo;
import org.zkoss.zk.ui.util.Composer;
import org.zkoss.zk.ui.util.ComposerExt;
import org.zkoss.zk.ui.util.FullComposer;

public class ZssappCtrl implements FullComposer, ComposerExt, Composer {

	@Override
	public ComponentInfo doBeforeCompose(Page page, Component parent,
			ComponentInfo compInfo) throws Exception {
		System.out.println("doBeforeCompose: " + parent);
		return compInfo;
	}

	@Override
	public void doBeforeComposeChildren(Component comp) throws Exception {
		System.out.println("doBeforeComposeChildren: " + comp);
		System.out.println("child");
	}

	@Override
	public boolean doCatch(Throwable ex) throws Exception {
		return false;
	}

	@Override
	public void doFinally() throws Exception {
	}

	@Override
	public void doAfterCompose(Component comp) throws Exception {
		System.out.println("ZssappCtrl doAfterCompose: " + comp);
		
	}	
}