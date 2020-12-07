package xuly;

import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;

public class ReadAndWriteExcelFile {
	
	public List<GiamThi> getListGiamThi(XSSFSheet sheet){
		List<GiamThi> listGiamThi = new ArrayList<>();
		DataFormatter fmt = new DataFormatter();
		for (Row row : sheet) {
			String maGV = "", hoTen = "", ngaySinh = "", donVi = "";
			Cell cellID = row.getCell(1);
			if(cellID == null) break;
			if (cellID.getCellType() == CellType.NUMERIC) {
				maGV = fmt.formatCellValue(cellID);
			} else if (cellID.getCellType() == CellType.STRING) {
				maGV = cellID.getStringCellValue();
			}
			//System.out.println(maGV);
			cellID = row.getCell(2);
			if (cellID.getCellType() == CellType.NUMERIC) {
				hoTen = fmt.formatCellValue(cellID);
			} else if (cellID.getCellType() == CellType.STRING) {
				hoTen = cellID.getStringCellValue();
			}
			cellID = row.getCell(3);
			if (cellID.getCellType() == CellType.NUMERIC) {
				ngaySinh = fmt.formatCellValue(cellID);
			} else if (cellID.getCellType() == CellType.STRING) {
				ngaySinh = cellID.getStringCellValue();
			}
			cellID = row.getCell(4);
			if (cellID.getCellType() == CellType.NUMERIC) {
				donVi = fmt.formatCellValue(cellID);
			} else if (cellID.getCellType() == CellType.STRING) {
				donVi = cellID.getStringCellValue();
			}
			listGiamThi.add(new GiamThi(maGV, hoTen, ngaySinh, donVi,"",""));
			
		}
		listGiamThi.remove(0);
		return listGiamThi;
	}
	
	public List<PhongThi> getListPhongThi(XSSFSheet sheet){
		List<PhongThi> listPhongThi = new ArrayList<>();
		for (Row row : sheet) {
			String phongThi = "", ghiChu = "";
			DataFormatter fmt = new DataFormatter();
			Cell cellID = row.getCell(1);
			if (cellID.getCellType() == CellType.NUMERIC) {
				phongThi = fmt.formatCellValue(cellID);
			} else if (cellID.getCellType() == CellType.STRING) {
				phongThi = cellID.getStringCellValue();
			}
			cellID = row.getCell(2);
			if (cellID.getCellType() == CellType.NUMERIC) {
				ghiChu = String.valueOf(cellID.getNumericCellValue());
			} else if (cellID.getCellType() == CellType.STRING) {
				ghiChu = cellID.getStringCellValue();
			}
			
			listPhongThi.add(new PhongThi(phongThi, ghiChu));
		}

		listPhongThi.remove(0);
		return listPhongThi;
	}
}
