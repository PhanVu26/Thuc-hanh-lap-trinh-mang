package xuly;

import java.io.InputStream;
import java.io.ObjectOutputStream;
import java.net.ServerSocket;
import java.net.Socket;
import java.util.ArrayList;
import java.util.List;
import java.util.Random;

import javax.sound.midi.Soundbank;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Server {
	
	public Server() {
		try {
			boolean isLoaded = false;
			ReadAndWriteExcelFile readWriteFile = new ReadAndWriteExcelFile();
			ServerSocket serversocket = new ServerSocket(4000);
			System.out.println("Server Started");
			List<GiamThi> listGiamThi = null;
			List<PhongThi> listPhongThi = null;
			List<GiamThi> listLichThi = new ArrayList<>();
			while (true) {
				Socket socket = serversocket.accept();
				System.out.println(socket);
				InputStream ips = socket.getInputStream();
				XSSFWorkbook workbook = new XSSFWorkbook(ips);
				XSSFSheet sheet = workbook.getSheetAt(0);
				XSSFSheet sheet1 = workbook.getSheetAt(1);

				if(isLoaded == false) {	
					System.out.println("loaded");
					listGiamThi = readWriteFile.getListGiamThi(sheet);
					listPhongThi = readWriteFile.getListPhongThi(sheet1);
					
					listLichThi = xepLich(listGiamThi, listPhongThi);
					isLoaded = true;				
				}else {
					listLichThi = xepLich(listGiamThi, listPhongThi);
				}
				System.out.println("hello");
				ObjectOutputStream oos = new ObjectOutputStream(socket.getOutputStream());
				oos.writeObject(listLichThi);				
				ips.close();
				socket.close();
				workbook.close();
			}

		} catch (Exception e) {
			System.out.println("Gặp lỗi khi chạy Server, thử kiểm tra lại cổng kết nối !");
			System.out.println(e);
		}
	}
	
	public List<GiamThi> xepLich(List<GiamThi> listGiamThi, List<PhongThi> listPhongThi) {
		List<GiamThi> listLichThi = new ArrayList<>();
		List<GiamThi> GiamThiChuaCoiThi = new ArrayList<>();
		
		for(int i = 0; i< listGiamThi.size(); i++) {
			GiamThiChuaCoiThi.add(listGiamThi.get(i));
		}
		GiamThi gt = new GiamThi();
		PhongThi pt = new PhongThi();
		//System.out.println("sizze" + listGiamThi.size() + listPhongThi.size());
		int soGT = 0;
		int k = -1;
		for(int i = 0; i< listPhongThi.size(); i++) {
			pt = listPhongThi.get(i);
			soGT = 0;
			for(int j= 0; j < GiamThiChuaCoiThi.size(); j++) {
				gt = GiamThiChuaCoiThi.get(j);
				//System.out.println(gt.getHoTen());
				//System.out.println(ktPhongDaXemThi(gt, pt));
				if(ktPhongDaXemThi(gt, pt) == false) {
					gt.getPhongDaCoi().add(pt);
					gt.setChucVu("Giám Thị " + (soGT + 1));
					gt.setPhongThi(pt.getPhongThi());
					soGT ++;
					k = j;
					listLichThi.add(gt);
					GiamThiChuaCoiThi.remove(j);
				}
				
				if(soGT == 2) break;
				
			}
		}
		
		int viTriGiamThiChuaXet = listLichThi.size() + 2;
		int index = 0;
		for(int i = viTriGiamThiChuaXet ; i < listGiamThi.size(); i++) {
			int soGTConLai = listGiamThi.size() - viTriGiamThiChuaXet;
			System.out.println(soGTConLai);
			int soPhong = soGTConLai / listPhongThi.size();
			System.out.println(soPhong);
			GiamThi gt1 = listGiamThi.get(i);
			gt1.setChucVu("Giám Thị hành lan");
			
			String phongGiamSat = "";
			int count = 0;
			for(int j = index; j < listPhongThi.size(); j++) {
				phongGiamSat += listPhongThi.get(j).getPhongThi() + ",";
				index = j;
				if(index == listPhongThi.size()) {
					phongGiamSat = "";
				}
				count ++;
				if(count == soPhong) {
					break;
				}
			}
			gt1.setPhongThi(phongGiamSat);
			listLichThi.add(gt1);
		}
		return listLichThi;
	}

	public boolean ktPhongDaXemThi(GiamThi gt, PhongThi phongThi) {
		for (PhongThi pt : gt.getPhongDaCoi()) {
			if(pt.getPhongThi().equals(phongThi.getPhongThi())) {
				System.out.println("aaaaa");
				return true;
			}
		}		
		return false;	
	}

	public static void main(String[] args) {
		new Server();
		
	}
}
