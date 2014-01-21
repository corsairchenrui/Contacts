import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashSet;
import java.util.Set;

import javax.imageio.stream.FileImageInputStream;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.util.CellRangeAddress;

public final class convertor1 {
	protected File f;
	protected HSSFWorkbook wb = null;
	protected FileOutputStream fos = null;
	protected HSSFSheet s;
	protected String header;
	private Set<contact> contacts;
	public convertor1(File f){
		this.f = f;
		header = f.getName().split("\\.")[0].trim();
		contacts = new HashSet<contact>();
		
	}
	public void load(){
		try {
			wb = new HSSFWorkbook(new POIFSFileSystem(new FileInputStream(f)));
			s = wb.getSheetAt(0);
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
	@Deprecated
	public void splitV(int r,int c){
		CellRangeAddress cra;
		for(int i=0;i<s.getNumMergedRegions();i++){
			 if((cra = s.getMergedRegion(i)).isInRange(r, c)){
				 break;
			 }
		}
		s.shiftRows(r + 1, s.getLastRowNum(), 1, true, false);
		for(int j = 0;j<9;j++){
			if(j == c)continue;
			Cell tmp = s.getRow(r).getCell(j);
			String str = tmp.getStringCellValue();
			CellRangeAddress cra1 = new CellRangeAddress(r, r+1, j, j);
			s.addMergedRegion(cra1);
			tmp.setCellValue(str);
		}
	}
	public void proceed(){
		for(int i=1;i<s.getLastRowNum();i++){
			for(int index = 2;index<8;){
				Cell a = s.getRow(i).getCell(index++);
				Cell b = s.getRow(i).getCell(index++);
				if(a.getStringCellValue().trim().equals(""))continue;
				contact[] c = analyze(a,b);
				if(c != null)
					for(contact tmp:c){
						boolean flag=true;
						for(contact item:contacts)
							if(item.equals(tmp)){
								item.category.addAll(tmp.category);
								flag=false;
								break;
							}
						if(flag)
							contacts.add(tmp);
					}
						
			}
		}
		for(contact c:contacts){
			System.err.println(c.lname+c.fname);
			System.err.println(c.tel);
			System.err.println(c.mobile);
			System.err.println(c.category);
		}
		
	}
	private contact[] analyze(Cell a, Cell b) {
		contact[] ret;
		String[] names = a.getStringCellValue().split("\n");
		String[] tels = b.getStringCellValue().split("\n");
		ret = new contact[names.length];
		int offset = 0;
		int stepping = tels.length/names.length;
		for(int i = 0;i<names.length;i++){
			String[] tel;
			if(i == names.length -1 ){
				tel = new String[tels.length - offset];
			}else{
				tel = new String[stepping];
			}
			for(int j = 0;j<tel.length;j++)
				tel[j] = tels[offset + j];
			contact tmp = new contact();
			String name = names[i].trim();
			if(name.length()<4){
				tmp.lname = name.substring(0, 1);
				tmp.fname = name.substring(1);
			}else{
				tmp.lname = name.substring(0, 2);
				tmp.fname = name.substring(2);
			}
			for(String t:tel){
				t = t.replaceAll("/", "").replaceAll("-", "").replaceAll(" ", "").trim();
				if(t.startsWith("1"))
					tmp.mobile = t;
				else if(t.startsWith("0"))
					tmp.tel = t;
				else
					try {
						throw new Exception("terrible");
					} catch (Exception e) {
						e.printStackTrace();
					}
			}
			tmp.category.add(this.header+a.getSheet().getRow(a.getRowIndex()).getCell(1).getStringCellValue());
			ret[i]=tmp;
			offset += tel.length;
		}
		
		return ret;
	}
	public void target(File out){
		try {
			fos = new FileOutputStream(out);
			wb.write(fos);
			fos.close();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
	public static final void step1(File dir){
		File dir1 = new File(dir,"1");
		for(File tmp:dir.listFiles()){
			if(tmp.isDirectory())continue;
			convertor1 c = new convertor1(tmp);
			c.load();
			
			c.target(new File(dir1,tmp.getName()));
		}
	}
	/**
	 * @param args
	 */
	public static void main(String[] args) {
		File dir = new File("D:\\contacts\\");
		File f = new File(dir, "test.et");
//		step1(dir);
		convertor1 a= new convertor1(f);
		a.load();
		a.proceed();
//		a.target(new File(f.getAbsoluteFile()+".et"));
	}

}
