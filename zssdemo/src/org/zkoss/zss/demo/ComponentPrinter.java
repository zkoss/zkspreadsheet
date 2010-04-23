package org.zkoss.zss.demo;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.PrintStream;
import java.util.Collection;
import java.util.Iterator;
import org.zkoss.zk.ui.*;
import org.zkoss.zk.ui.metainfo.Annotation;
import org.zkoss.zk.ui.sys.ComponentCtrl;
/**
 *
 * @author Dennis.Chen
 */
public class ComponentPrinter {

    
    static public void print(Component comp){
        print(comp,System.out);
    }
    static public void print(Component comp,PrintStream out){
        recursivePrint(comp,out,0);
    }
    static public void print(Desktop desktop){
        print(desktop,System.out);   
    }
    static public void print(Desktop desktop,StringBuffer sb){
        try {
            ByteArrayOutputStream os = new ByteArrayOutputStream();
            print(desktop, new PrintStream(os, true));
            os.close();
            sb.append(new String(os.toByteArray()));
        } catch (IOException ex) {
            ex.printStackTrace();
        }
        
    }
    static public void print(Desktop desktop,PrintStream out){
        Collection col = desktop.getPages();
        out.println("Desktop->"+desktop.getId());
        for(Iterator iter = col.iterator();iter.hasNext();){
            out.println("");
            print((Page)iter.next(),out);
        } 
    }
    static public void print(Page page){
        print(page,System.out);   
    }
    static public void print(Page page,PrintStream out){
        out.println("Page->"+page.getId());
        Collection col = page.getRoots();
        for(Iterator iter = col.iterator();iter.hasNext();){
            print((Component)iter.next(),out);
        }
    }
    
	static private void recursivePrint(Component comp, PrintStream out, int level)
	{
		StringBuffer sb = new StringBuffer();
		for(int i=level;i>0;i--)out.print(" |");
		dump(sb,comp);
		out.println(sb);
		for(Iterator itor = comp.getChildren().iterator();itor.hasNext();)
			recursivePrint((Component) itor.next(),out,level+1);
	}
	private static void dump(StringBuffer sb, Component comp) 
	{
        if(comp instanceof ComponentPrinter) return;
		String srClss=comp.getClass().toString();
		srClss = srClss.substring(srClss.lastIndexOf(".")+1);
		ComponentCtrl compCtrl = (ComponentCtrl)comp;
		
		//sb.append("->"+srClss+"-"+comp.getId() +" "+comp.toString());
		sb.append("->"+comp.toString());
        
        //print annotation;
        if(true) return;
		sb.append("=").append(compCtrl.getAnnotations().size()+" : ")
			.append(compCtrl.getAnnotatedProperties().size());
		String prop;
		for (Iterator it = compCtrl.getAnnotations().iterator(); it.hasNext();) {
			Annotation annot = (Annotation) it.next();
			sb.append(" self").append(annot);
		}
		for (Iterator it = compCtrl.getAnnotatedProperties().iterator(); it.hasNext();) {
			prop = (String) it.next();
			sb.append(" ").append(prop)
			.append(compCtrl.getAnnotations(prop));
		}
	}
}
