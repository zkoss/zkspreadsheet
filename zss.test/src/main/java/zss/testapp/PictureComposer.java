package zss.testapp;

import java.io.IOException;

import org.zkoss.bind.annotation.Init;
import org.zkoss.image.AImage;
import org.zkoss.zk.ui.*;
import org.zkoss.zk.ui.select.SelectorComposer;
import org.zkoss.zk.ui.select.annotation.*;
import org.zkoss.zss.model.*;
import org.zkoss.zss.range.SRanges;
import org.zkoss.zss.ui.Spreadsheet;
import org.zkoss.zul.*;

/**
 * Demonstrate picture related API usage.
 * @author Hawk
 * 
 */
@SuppressWarnings("serial")
public class PictureComposer extends SelectorComposer<Component> {

	@Wire
	private Intbox toRowBox;
	@Wire
	private Intbox toColumnBox;
	@Wire
	private Spreadsheet ss;
	@Wire
	private Listbox pictureListbox;
	
	private ListModelList<SPicture> pictureList = new ListModelList<SPicture>();


	@Init
	public void init(){
		refreshPictureList();
	}
	
	@Listen("onClick = #addButton")
	public void add() {
		try{ 
			AImage image = new AImage(WebApps.getCurrent().getResource("/range/zklogo.png"));
			ViewAnchor anchor = new ViewAnchor(ss.getSelection().getRow(), ss.getSelection().getColumn(), image.getWidth(), image.getHeight());

			SRanges.range(ss.getSelectedSSheet()).addPicture(anchor, image.getByteData(), SPicture.Format.PNG);
			refreshPictureList();
		}catch(IOException e){
			System.out.println("cannot add a picture for "+ e);
		}
	}
	
	
	@Listen("onClick = #deleteButton")
	public void delete() {
		if (pictureListbox.getSelectedItem() != null){
			SRanges.range(ss.getSelectedSSheet()).deletePicture((SPicture)pictureListbox.getSelectedItem().getValue());
			refreshPictureList();
		}
	}
	
	@Listen("onClick = #moveButton")
	public void move() {
		if (pictureListbox.getSelectedItem() != null){
			
			SPicture picture = (SPicture) pictureListbox.getSelectedItem().getValue();
			ViewAnchor anchor = new ViewAnchor(toRowBox.getValue(), toColumnBox.getValue(), picture.getAnchor().getWidth(), picture.getAnchor().getHeight());
			SRanges.range(ss.getSelectedSSheet()).movePicture(picture, anchor);
			refreshPictureList();
		}
	}
	
	private void refreshPictureList(){
		pictureList.clear();
		pictureList.addAll(ss.getSelectedSSheet().getPictures());
		pictureListbox.setModel(pictureList);
	}
	
//	@Listen("onClick = #addButton")
//	public void addByUtil() {
//		Range selection = Ranges.range(ss.getSelectedSheet(), ss.getSelection());
//		try{
//			SheetOperationUtil.addPicture(selection,
//				new AImage(WebApps.getCurrent().getResource("/zklogo.png")));
//			refreshPictureList();
//		}catch(IOException e){
//			System.out.println("cannot add a picture for "+ e);
//		}
//	}
//	
//	@Listen("onClick = #deleteButton")
//	public void deleteByUtil() {
//		if (pictureListbox.getSelectedItem() != null){
//			SheetOperationUtil.deletePicture(Ranges.range(ss.getSelectedSheet()),
//					(Picture)pictureListbox.getSelectedItem().getValue());
//			refreshPictureList();
//		}
//	}
//	
//	@Listen("onClick = #moveButton")
//	public void moveByUtil(){
//		if (pictureListbox.getSelectedItem() != null){
//			SheetOperationUtil.movePicture(Ranges.range(ss.getSelectedSheet()),
//					(Picture)pictureListbox.getSelectedItem().getValue(),
//					toRowBox.getValue(), toColumnBox.getValue());
//			refreshPictureList();
//		}
//	}
//	
//	@Listen("onClick = #exportButton")
//	public void export() throws IOException {
//		Exporter excelExporter = Exporters.getExporter("excel");
//		File file = new File("exported.xlsx");
//		FileOutputStream fos = new FileOutputStream(file);
//		excelExporter.export(ss.getBook(), fos);
//		Filedownload.save(file, "application/excel");
//	}
}
