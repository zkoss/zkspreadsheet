package org.zkoss.zss.ngmodel.impl;

import java.io.Serializable;

import org.zkoss.zss.ngmodel.NName;
import org.zkoss.zss.ngmodel.sys.dependency.Ref;

public abstract class NameAdv implements NName,LinkedModelObject,FormulaContent,Serializable{

	abstract void setName(String newname);

	abstract BookAdv getBook();

	

}
