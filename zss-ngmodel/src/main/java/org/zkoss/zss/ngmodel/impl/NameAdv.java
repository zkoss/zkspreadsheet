package org.zkoss.zss.ngmodel.impl;

import java.io.Serializable;

import org.zkoss.zss.ngmodel.NName;

public abstract class NameAdv implements NName,LinkedModelObject,Serializable{

	private static final long serialVersionUID = 1L;

	abstract void setName(String newname);

	abstract public BookAdv getBook();

	

}
