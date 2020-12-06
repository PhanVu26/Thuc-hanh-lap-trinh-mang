package xuly;

import java.io.Serializable;

public class PhongThi implements Serializable{
	String phongThi;
	String ghiChu;
	
	public PhongThi() {
		// TODO Auto-generated constructor stub
	}

	public PhongThi(String phongThi, String ghiChu) {
		super();
		this.phongThi = phongThi;
		this.ghiChu = ghiChu;
	}

	public String getPhongThi() {
		return phongThi;
	}

	public void setPhongThi(String phongThi) {
		this.phongThi = phongThi;
	}

	public String getGhiChu() {
		return ghiChu;
	}

	public void setGhiChu(String ghiChu) {
		this.ghiChu = ghiChu;
	}
	
	
}
