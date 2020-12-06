package xuly;

import java.io.IOException;
import java.util.List;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Test {

	public static void main(String[] args) {
		ReadAndWriteExcelFile readWriteFile = new ReadAndWriteExcelFile();
		
		XSSFWorkbook workbook;
		try {
			workbook = new XSSFWorkbook("D:/java/Thuc-hanh-lap-trinh-mang/input-data/Danh-sach-can-bo-coi-thi.xlsx");
			XSSFSheet sheet = workbook.getSheetAt(0);
			//System.out.println(sheet);
			XSSFSheet sheet1 = workbook.getSheetAt(1);
			//System.out.println(sheet1);
			List<GiamThi> listGiamThi = readWriteFile.getListGiamThi(sheet);
			List<PhongThi> listPhongThi = readWriteFile.getListPhongThi(sheet1);
			int c = 0;
			for (PhongThi pt : listPhongThi) {
				System.out.println(pt.getPhongThi());
				c ++;
				if(c==10) break;
			}
			c = 0;
			for (GiamThi pt : listGiamThi) {
				System.out.println(pt.getHoTen());
				c ++;
				if(c==10) break;
			}
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
		
	}

}
