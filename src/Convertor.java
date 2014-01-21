import java.io.File;
import java.io.IOException;

import jxl.Cell;
import jxl.Range;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;



public class Convertor {
	public static final Range getMergedCell(Sheet s,String cellName){
		if(null != s && null != cellName && !cellName.isEmpty() && cellName.trim().length() == 2){
			cellName = cellName.toUpperCase();
			Cell c = null;
			try{
				c = s.getCell(cellName);
			} catch(ArrayIndexOutOfBoundsException e){
			}
			if(null != c){
				Range[] ranges = s.getMergedCells();
				for(Range r:ranges){
					String[] range = r.toString().split("-");
					if((range[0].charAt(0) <= cellName.charAt(0) && range[0].charAt(1) <= cellName.charAt(1)) && 
							(range[1].charAt(0) >= cellName.charAt(0) && range[1].charAt(1) >= cellName.charAt(1))){
						return r;
					}
				}
			}
		}
		return null;
	}
	/**
	 * @param args
	 */
	public static void main(String[] args) {
		// TODO Auto-generated method stub
//		File f = new File("C:\\Users\\abc\\Desktop\\分行应用系统保障人员通讯录\\附件1：湖南分行应用系统保障人员通讯录.et");
		File f = new File("C:\\Users\\abc\\Desktop\\test.xls");
		try {
			Workbook wb = Workbook.getWorkbook(f);
			Sheet s = wb.getSheet(0);
			WritableWorkbook wwb = Workbook.createWorkbook(new File(f.getAbsolutePath()), wb);
			Object o = wwb.getSheet(0).getCell("A2").toString();
//			Object o = getMergedCell(s, "a1").getTopLeft().getContents();
			System.err.print(o);
//			wwb.write();
			wwb.close();
			wb.close();
		} catch (BiffException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (WriteException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
	}

}
