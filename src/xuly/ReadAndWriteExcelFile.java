package xuly;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

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
	
	public void writeFile(List<GiamThi> listPT) {
XSSFWorkbook workbook = new XSSFWorkbook();
		
		int GTPerSheet = 20;
		int sheetNumber = listPT.size() / GTPerSheet;
		
		
		XSSFSheet newSheet1 = workbook.createSheet("Phân công tổng hợp");
		Row row01 = newSheet1.createRow(0);
		Cell cell01 = row01.createCell(3);
		//Cell cell1 = row0.createCell(4);
		newSheet1.addMergedRegion(new CellRangeAddress(0, 0, 3, 6));
		cell01.setCellValue("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM");
		
		CellStyle style1 = workbook.createCellStyle();
		XSSFFont font = workbook.createFont();
		font.setFontHeightInPoints((short) 14);
		style1.setFont(font);
		style1.setAlignment(HorizontalAlignment.CENTER);
		cell01.setCellStyle(style1);
		
		Row rowTitle1 = newSheet1.createRow(1);
		Cell cellTitle1 = rowTitle1.createCell(3);
		newSheet1.addMergedRegion(new CellRangeAddress(1, 1, 3, 6));
		cellTitle1.setCellValue("Độc lập - Tự do - Hạnh phúc");
		
		CellStyle style11 = workbook.createCellStyle();
		XSSFFont font11 = workbook.createFont();
		font11.setFontHeightInPoints((short) 13);
		font11.setUnderline(HSSFFont.U_SINGLE);
		style1.setFont(font11);
		style1.setAlignment(HorizontalAlignment.CENTER);
		cellTitle1.setCellStyle(style1);
		
		
		Row rowTitle21 = newSheet1.createRow(3);
		Cell cellTitle21 = rowTitle21.createCell(3);
		newSheet1.addMergedRegion(new CellRangeAddress(3, 3, 3, 6));
		cellTitle21.setCellValue("DANH SÁCH PHÂN CÔNG GIÁM THỊ COI THI");
		XSSFFont font2 = workbook.createFont();
		
		CellStyle style21 = workbook.createCellStyle();
		font2.setFontHeightInPoints((short) 16);
		font2.setBold(true);
		style21.setFont(font2);
		style21.setAlignment(HorizontalAlignment.CENTER);
		cellTitle21.setCellStyle(style21);
		
		CellStyle styleheader1 = workbook.createCellStyle();
		XSSFFont fontheader1 = workbook.createFont();
		fontheader1.setBold(true);
		styleheader1.setFont(fontheader1);
		styleheader1.setBorderBottom(BorderStyle.MEDIUM);
		styleheader1.setBorderLeft(BorderStyle.MEDIUM);
		styleheader1.setBorderRight(BorderStyle.MEDIUM);
		styleheader1.setBorderTop(BorderStyle.MEDIUM);
		
		Row headRow1 = newSheet1.createRow(5);
		Cell col11 = headRow1.createCell(1);
		col11.setCellValue("STT");
		col11.setCellStyle(styleheader1);
		
		Cell col21 = headRow1.createCell(2);
		col21.setCellValue("Mã số GV");
		col21.setCellStyle(styleheader1);
		
		Cell col31 = headRow1.createCell(3);
		col31.setCellValue("Họ và tên");
		col31.setCellStyle(styleheader1);
		
		Cell col41 = headRow1.createCell(4);
		col41.setCellValue("Ngày Sinh");
		col41.setCellStyle(styleheader1);
		
		Cell col51 = headRow1.createCell(5);
		col51.setCellValue("Đơn vị");
		col51.setCellStyle(styleheader1);
		
		Cell col61 = headRow1.createCell(6);
		col61.setCellValue("Phòng thi");
		col61.setCellStyle(styleheader1);
		
		Cell col71 = headRow1.createCell(7);
		col71.setCellValue("Chức vụ");
		col71.setCellStyle(styleheader1);
		
		
		XSSFFont fontbody1 = workbook.createFont();
		CellStyle stylebody1 = workbook.createCellStyle();
		stylebody1.setFont(fontbody1);
		stylebody1.setBorderBottom(BorderStyle.MEDIUM);
		stylebody1.setBorderLeft(BorderStyle.MEDIUM);
		stylebody1.setBorderRight(BorderStyle.MEDIUM);
		stylebody1.setBorderTop(BorderStyle.MEDIUM);
		
		int numRow =6;
		int count = 0;
		for (int j = 0; j < listPT.size(); j++) {
			GiamThi pt = listPT.get(j);
			Row row = newSheet1.createRow(numRow);
			Cell cell1 = row.createCell(1);
			cell1.setCellValue(count+1);
			cell1.setCellStyle(stylebody1);
			
			Cell cell2 = row.createCell(2);
			cell2.setCellValue(pt.getMaGV());
			cell2.setCellStyle(stylebody1);
			
			Cell cell3 = row.createCell(3);
			cell3.setCellValue(pt.getHoTen());
			cell3.setCellStyle(stylebody1);
			
			
			Cell cell4 = row.createCell(4);
			cell4.setCellValue(pt.getNgaySinh());
			cell4.setCellStyle(stylebody1);
			
			Cell cell5 = row.createCell(5);
			cell5.setCellValue(pt.getDonVi());
			cell5.setCellStyle(stylebody1);
			
			Cell cell6 = row.createCell(6);
			cell6.setCellValue(pt.getPhongThi());
			cell6.setCellStyle(stylebody1);
			
			Cell cell7 = row.createCell(7);
			cell7.setCellValue(pt.getChucVu());
			cell7.setCellStyle(stylebody1);
			
			count++;
			numRow++;
		}
		newSheet1.autoSizeColumn(0);
		newSheet1.autoSizeColumn(1);
		newSheet1.autoSizeColumn(2);
		newSheet1.autoSizeColumn(3);
		newSheet1.autoSizeColumn(4);
		newSheet1.autoSizeColumn(5);
		newSheet1.autoSizeColumn(6);
		newSheet1.autoSizeColumn(7);
		
		
		Row rowSign1 = newSheet1.createRow(listPT.size() + 7);
		Cell cellSign1 = rowSign1.createCell(6);
		cellSign1.setCellValue("Hội đồng thi");
		
		Row rowSign2 = newSheet1.createRow(listPT.size() + 9);
		Cell cellSign2 = rowSign2.createCell(6);
		cellSign2.setCellValue("Phan Văn Vũ");
		
		System.out.println("print");
	
		
		
		for (int i = 0; i < sheetNumber; i++) {
		
			XSSFSheet newSheet = workbook.createSheet("Phân công " + (i+1));
			Row row0 = newSheet.createRow(0);
			Cell cell0 = row0.createCell(3);
			//Cell cell1 = row0.createCell(4);
			newSheet.addMergedRegion(new CellRangeAddress(0, 0, 3, 6));
			cell0.setCellValue("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM");
			
			CellStyle style = workbook.createCellStyle();
			XSSFFont font1 = workbook.createFont();
			font1.setFontHeightInPoints((short) 14);
			style.setFont(font1);
			style.setAlignment(HorizontalAlignment.CENTER);
			cell0.setCellStyle(style);
			
			Row rowTitle11 = newSheet.createRow(1);
			Cell cellTitle11 = rowTitle11.createCell(3);
			newSheet.addMergedRegion(new CellRangeAddress(1, 1, 3, 6));
			cellTitle11.setCellValue("Độc lập - Tự do - Hạnh phúc");
			
			CellStyle style111 = workbook.createCellStyle();
			XSSFFont font111 = workbook.createFont();
			font111.setFontHeightInPoints((short) 13);
			font111.setUnderline(HSSFFont.U_SINGLE);
			style111.setFont(font111);
			style111.setAlignment(HorizontalAlignment.CENTER);
			cellTitle11.setCellStyle(style111);
			
			
			Row rowTitle2 = newSheet.createRow(3);
			Cell cellTitle2 = rowTitle2.createCell(3);
			newSheet.addMergedRegion(new CellRangeAddress(3, 3, 3, 6));
			cellTitle2.setCellValue("DANH SÁCH PHÂN CÔNG COI THI");
			XSSFFont font21 = workbook.createFont();
			
			CellStyle style2 = workbook.createCellStyle();
			font21.setFontHeightInPoints((short) 16);
			font21.setBold(true);
			style2.setFont(font21);
			style2.setAlignment(HorizontalAlignment.CENTER);
			cellTitle2.setCellStyle(style2);
			
			CellStyle styleheader = workbook.createCellStyle();
			XSSFFont fontheader = workbook.createFont();
			fontheader.setBold(true);
			styleheader.setFont(fontheader);
			styleheader.setBorderBottom(BorderStyle.MEDIUM);
			styleheader.setBorderLeft(BorderStyle.MEDIUM);
			styleheader.setBorderRight(BorderStyle.MEDIUM);
			styleheader.setBorderTop(BorderStyle.MEDIUM);
			
			Row headRow = newSheet.createRow(5);
			Cell col1 = headRow.createCell(1);
			col1.setCellValue("STT");
			col1.setCellStyle(styleheader);
			
			Cell col2 = headRow.createCell(2);
			col2.setCellValue("Mã số GV");
			col2.setCellStyle(styleheader);
			
			Cell col3 = headRow.createCell(3);
			col3.setCellValue("Họ và tên");
			col3.setCellStyle(styleheader);
			
			Cell col4 = headRow.createCell(4);
			col4.setCellValue("Ngày Sinh");
			col4.setCellStyle(styleheader);
			
			Cell col5 = headRow.createCell(5);
			col5.setCellValue("Đơn vị");
			col5.setCellStyle(styleheader);
			
			Cell col6 = headRow.createCell(6);
			col6.setCellValue("Phòng thi");
			col6.setCellStyle(styleheader);
			
			Cell col7 = headRow.createCell(7);
			col7.setCellValue("Chức vụ");
			col7.setCellStyle(styleheader);
			

			
			XSSFFont fontbody = workbook.createFont();
			CellStyle stylebody = workbook.createCellStyle();
			stylebody.setFont(fontbody);
			stylebody.setBorderBottom(BorderStyle.MEDIUM);
			stylebody.setBorderLeft(BorderStyle.MEDIUM);
			stylebody.setBorderRight(BorderStyle.MEDIUM);
			stylebody.setBorderTop(BorderStyle.MEDIUM);
			
			int numRow1 =6;
			int count1 = 0;
			for (int j = i * GTPerSheet; j <= i * GTPerSheet + GTPerSheet - 1; j++) {
				GiamThi pt = listPT.get(j);
				Row row = newSheet.createRow(numRow1);
				Cell cell1 = row.createCell(1);
				cell1.setCellValue(count1+1);
				cell1.setCellStyle(stylebody);
				
				Cell cell2 = row.createCell(2);
				cell2.setCellValue(pt.getMaGV());
				cell2.setCellStyle(stylebody);
				
				Cell cell3 = row.createCell(3);
				cell3.setCellValue(pt.getHoTen());
				cell3.setCellStyle(stylebody);
				
				
				Cell cell4 = row.createCell(4);
				cell4.setCellValue(pt.getNgaySinh());
				cell4.setCellStyle(stylebody);
				
				Cell cell5 = row.createCell(5);
				cell5.setCellValue(pt.getDonVi());
				cell5.setCellStyle(stylebody);
				
				Cell cell6 = row.createCell(6);
				cell6.setCellValue(pt.getPhongThi());
				cell6.setCellStyle(stylebody);
				
				Cell cell7 = row.createCell(7);
				cell7.setCellValue(pt.getChucVu());
				cell7.setCellStyle(stylebody);

				
				
				
				count1++;
				numRow1++;
			}
			newSheet.autoSizeColumn(0);
			newSheet.autoSizeColumn(1);
			newSheet.autoSizeColumn(2);
			newSheet.autoSizeColumn(3);
			newSheet.autoSizeColumn(4);
			newSheet.autoSizeColumn(5);
			newSheet.autoSizeColumn(6);
			newSheet.autoSizeColumn(7);
			
			
			Row rowSign11 = newSheet.createRow(27);
			Cell cellSign11 = rowSign11.createCell(6);
			cellSign11.setCellValue("Hội đồng thi");
			
			Row rowSign21 = newSheet.createRow(29);
			Cell cellSign21 = rowSign21.createCell(6);
			cellSign21.setCellValue("Phan Văn Vũ");
			
			System.out.println("Đang ghi file");
		}
		
		System.out.println("Hoàn thành");
			
		File fileout = new File ("./output-data/phanCong.xlsx");
		try {
			FileOutputStream out = new FileOutputStream(fileout);
			workbook.write(out);
			out.close();
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
}
