package xuly;

import java.awt.Color;
import java.awt.Dimension;
import java.awt.FileDialog;
import java.awt.Frame;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.BufferedInputStream;
import java.io.BufferedOutputStream;
import java.io.DataInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FilenameFilter;
import java.io.IOException;
import java.io.InputStream;
import java.io.ObjectInputStream;
import java.io.OutputStream;
import java.net.Socket;
import java.util.ArrayList;
import java.util.List;
import java.util.Vector;

import javax.swing.JButton;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JPanel;
import javax.swing.JScrollPane;
import javax.swing.JTable;
import javax.swing.JTextField;
import javax.swing.table.DefaultTableModel;
import javax.swing.table.TableColumn;

import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


import java.awt.SystemColor;
import javax.swing.JLabel;

public class Client extends JFrame implements ActionListener {

//components
	JButton btnChooseFile;
	FileDialog dialog;
	String pathFile;
	JTextField path ;
	JButton btnSendFile;
	JPanel result ;
	JScrollPane sp ;
	JTable table ;
	DefaultTableModel tableModel = new DefaultTableModel();

//	--------------------------------

	private JLabel lblDanhSchPhn;
//	static Socket sock;

	public static void main(String[] argv) throws Exception {
		new Client();
	}

	public Client() {
		getContentPane().setBackground(SystemColor.activeCaption);
		setBackground(SystemColor.activeCaption);
		init();
		getWidget();
	}

	
	public void init() {
		this.btnChooseFile = new JButton("Choose file");
		btnChooseFile.addActionListener(this);
		this.btnSendFile = new JButton("Get result");
		btnSendFile.addActionListener(this);
		dialog = new FileDialog((Frame) null, "Select File to Open");
		dialog.setMode(FileDialog.LOAD);
		dialog.setFilenameFilter(new FilenameFilter(){
            @Override
            public boolean accept(File dir, String name) {
                return name.endsWith(".xlsx");
            }
		});
		table = new JTable();
		table.setBounds(0, 60, 800, 540);
		sp = new JScrollPane(table);
		sp.setBounds(10, 60, 974, 542);
		table.setModel(tableModel);
		getContentPane().add(sp);
		path = new JTextField();
		result = new JPanel();
	}

	public void getWidget() {
		this.setTitle("Phân chia giám thị coi thi");
		this.setLocation(200, 20);
		this.setSize(1000, 700);
		this.setDefaultCloseOperation(EXIT_ON_CLOSE);
		getContentPane().setLayout(null);
		this.setResizable(false);
		path.setBounds(10, 613 , 628, 29);
		btnChooseFile.setBounds(648, 613, 100, 29);
		btnSendFile.setBounds(851, 613, 100, 29);
		getContentPane().add(btnChooseFile);
		getContentPane().add(btnSendFile);
		getContentPane().add(path);
		
		lblDanhSchPhn = new JLabel("PHẦN MỀM PHÂN CÔNG GIÁM THỊ COI THI");
		lblDanhSchPhn.setFont(new java.awt.Font("Tahoma", java.awt.Font.BOLD, 15));
		lblDanhSchPhn.setBounds(299, 11, 345, 38);
		getContentPane().add(lblDanhSchPhn);
		this.setVisible(true);

	}

	public void UpdateTable (List<GiamThi> listGT) {
		try {
			
			tableModel.getDataVector().removeAllElements();
			String[] columnNames = { "STT", "Mã Giáo viên", "Họ và tên", "Ngày Sinh", "Đơn vị", "Phòng thi", "Chức vụ"}; 
			tableModel.setColumnIdentifiers(columnNames);
			
			
			int tt = 1;
			for (GiamThi gt :listGT ) {
				String rows[] = new String[] {tt+"" , gt.getMaGV() ,
						gt.getHoTen(), gt.getNgaySinh(), 
						gt.getDonVi(), gt.getPhongThi(), 
						gt.getChucVu()};
				tableModel.addRow(rows);
				tt++;
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
	public boolean sendFile(String urlFile)  {
		File file = new File(urlFile);
		byte[] byteArray = new byte[(int) file.length()];
		int arrayByteLength = byteArray.length;
		int length;
		try {
			Socket sock = new Socket("localhost", 4000);
			BufferedInputStream bis = new BufferedInputStream(new FileInputStream(file));
			bis.read(byteArray, 0, arrayByteLength);
			FileInputStream fis = new FileInputStream(file);
			OutputStream ops = sock.getOutputStream();
			ops.write(byteArray, 0, arrayByteLength);
			ObjectInputStream ois = new ObjectInputStream(sock.getInputStream());
			Object result = ois.readObject();
			
			List<GiamThi> listPT =  (ArrayList<GiamThi>) result;
			ReadAndWriteExcelFile readWriteFile = new ReadAndWriteExcelFile();
			//getFile(listPT);
			
			UpdateTable(listPT);
			readWriteFile.writeFile(listPT);
		} catch (FileNotFoundException e) {
			System.out.println("can not read this file !");
			e.printStackTrace();
		} catch (IOException e) {
			System.out.println("have error when get input stream from socket!");
			e.printStackTrace();
		} catch (ClassNotFoundException e) {
			System.out.println("Error");
		}

		return false;
	}

	@Override
	public void actionPerformed(ActionEvent e) {
		// TODO Auto-generated method stub
		if (e.getActionCommand().equals("Choose file")) {
			dialog.setVisible(true);
			String directory = dialog.getDirectory();
			String nameFile = dialog.getFile();
			pathFile = directory + nameFile;
			System.out.println(pathFile);
			path.setText(pathFile);
		}
		
		if (e.getActionCommand().equals("Get result")) {
			try {
				sendFile(pathFile);
			} catch (Exception e2) {
				System.out.println("File không tồn tại hoặc không đúng định dạng");
			}
			
			
		}
	}
	
	
}
