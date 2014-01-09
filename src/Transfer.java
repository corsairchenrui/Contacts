import java.io.File;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.FilenameFilter;
import java.io.IOException;
import java.io.OutputStreamWriter;
import java.sql.Statement;

import MySQL.MySQLEntity;
import MySQL.field;
import MySQL.DAO.DAOFactory;
import MySQL.DAO.DAOImpl;
import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;

public class Transfer {
	public static class Row extends MySQLEntity{
		@field(type="s")
		protected String name;
		@field(type="s")
		protected String fname;
		@field(type="s")
		protected String category;
		@field(type="s")
		protected String mobile;
		@field(type="s")
		protected String tel;
		
		public Row(String tablename) {
			super(tablename);
		}

		public String getName() {
			return name;
		}

		public void setName(String name) {
			this.name = name;
		}

		public String getFname() {
			return fname;
		}

		public void setFname(String fname) {
			this.fname = fname;
		}

		public String getCategory() {
			return category;
		}

		public void setCategory(String category) {
			this.category = category;
		}

		public String getMobile() {
			return mobile;
		}

		public void setMobile(String mobile) {
			this.mobile = mobile;
		}

		public String getTel() {
			return tel;
		}

		public void setTel(String tel) {
			this.tel = tel;
		}

		public static Row newRow(Cell[] cell) {
			if(null != cell){
				Row row = new Row("data1");
				row.name = cell[0].getContents().trim();
				row.fname = cell[1].getContents().trim();
				row.category = cell[4].getContents().trim();
				row.mobile = cell[2].getContents().trim();
				if(!cell[4].getContents().trim().isEmpty())
					row.tel = cell[3].getContents().trim().split("/")[0];
				else
					row.tel = cell[3].getContents().trim();
				return row;
			}
			return null;
		}
//		public static Row newRow(Cell[] cell) {
//			if(null != cell){
//				Row row = new Row("data1");
//				row.name = cell[1].getContents().trim();
//				row.fname = cell[3].getContents().trim();
//				row.category = cell[6].getContents().trim();
//				row.mobile = cell[40].getContents().trim();
//				if(!cell[44].getContents().trim().isEmpty())
//					row.tel = cell[44].getContents().trim().split("/")[0];
//				else
//					row.tel = cell[44].getContents().trim();
//				return row;
//			}
//			return null;
//		}
		
	}

	private static DAOImpl dao;
	private static File dir;
	private static FileWriter fw = null;
	private static OutputStreamWriter writer;
	private static void Import() throws BiffException, IOException{
		File[] xlsFiles = dir.listFiles(new FilenameFilter() {
			
			@Override
			public boolean accept(File arg0, String arg1) {
				if(arg1.trim().equals("数据中心通讯录.xls"))
					return true;
				return false;
			}
		});
		System.err.println(xlsFiles.length);
		for(File f:xlsFiles){
			Workbook workbook = Workbook.getWorkbook(f);
			parse(workbook.getSheet(0));
		}
	}
	private static void Export(){
		
	}
	public static void main(String[] args) {
		dao = DAOImpl.newDAO("contact");
		dir = new File("F:\\workspace\\vcf");
		try {
			writer = new OutputStreamWriter(new FileOutputStream(new File(dir, "数据中心.vcf")), "UTF-8");
			Import();
			Export();
			writer.close();
		} catch (BiffException | IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
	}
	private static void parse(Sheet sheet) throws IOException {
		for(int i = 0;i<sheet.getRows();i++){
			Row row=null;
try{
			row = Row.newRow(sheet.getRow(i));
}catch(ArrayIndexOutOfBoundsException e){
	e.printStackTrace();
}
			if(!row.getName().trim().isEmpty()){
//				dao.insert(row);
				writer.append("BEGIN:VCARD\nVERSION:3.0\n");
				writer.append("N:"+row.fname+";"+row.name+";;;\n");
				writer.append("FN:"+row.fname+row.name+"\n");
				writer.append("TEL;TYPE=CELL:"+row.mobile+"\n");
				writer.append("TEL;TYPE=WORK:"+row.tel+"\n");
				writer.append("CATEGORIES:"+row.category+"\n");
				writer.append("X-WDJ-STARRED:0\nEND:VCARD\n");
				writer.flush();
			}
		}
	}
//	private static void parse(Sheet sheet) throws IOException{
//		OutputStreamWriter bla = new OutputStreamWriter(new FileOutputStream(new File("F:/workspace/vcf/bla")), "UTF-8");
//		for(int i = 0;i<sheet.getRows();i++){
//			Cell[] cells = sheet.getRow(i);
//			if(cells[0].getContents().trim().isEmpty()||cells[1].getContents().trim().isEmpty()){
//				continue;
//			}
//			String s = cells[2].getContents().trim();
//			String a = s.split("/")[0].trim();
//			String b = s.split("/")[1].trim();
//			bla.append(b+a+"\n");
//		}
//		bla.flush();
//		bla.close();
//	}
}
