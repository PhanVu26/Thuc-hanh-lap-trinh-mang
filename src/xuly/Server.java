package xuly;

import java.io.InputStream;
import java.io.ObjectOutputStream;
import java.net.ServerSocket;
import java.net.Socket;
import java.util.List;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Server {
	
	public Server() {
		try {
			ReadAndWriteExcelFile readWriteFile = new ReadAndWriteExcelFile();
			ServerSocket serversocket = new ServerSocket(4000);
			System.out.println("Server Started");
			while (true) {
				Socket socket = serversocket.accept();
				InputStream ips = socket.getInputStream();
				XSSFWorkbook workbook = new XSSFWorkbook(ips);
				XSSFSheet sheet = workbook.getSheetAt(0);
				XSSFSheet sheet1 = workbook.getSheetAt(1);
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
				ObjectOutputStream oos = new ObjectOutputStream(socket.getOutputStream());
				oos.writeObject(listGiamThi);
				workbook.close();
				ips.close();
				socket.close();
			}

		} catch (Exception e) {
			System.out.println("Gặp lỗi khi chạy Server, thử kiểm tra lại cổng kết nối !");
			System.out.println(e);
		}
	}
	
	public static void main(String[] args) {
		new Server();
		
	}
}
