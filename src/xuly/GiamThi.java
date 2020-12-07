package xuly;

import java.io.Serializable;
import java.util.ArrayList;
import java.util.List;

public class GiamThi implements Serializable{
	String maGV;
	String hoTen;
	String ngaySinh;
	String donVi;
	String chucVu;
	String phongThi;
	List<PhongThi> phongDaCoi = new ArrayList<>();
	List<GiamThi> listGT = new ArrayList<>();
	
	public GiamThi(String maGV, String hoTen, String ngaySinh, String donVi, String chucVu, String phongThi) {
		super();
		this.maGV = maGV;
		this.hoTen = hoTen;
		this.ngaySinh = ngaySinh;
		this.donVi = donVi;
		this.chucVu = chucVu;
		this.phongThi = phongThi;
	}
	
	public List<PhongThi> getPhongDaCoi() {
		return phongDaCoi;
	}

	public void setPhongDaCoi(List<PhongThi> listPT) {
		this.phongDaCoi = listPT;
	}

	public List<GiamThi> getListGT() {
		return listGT;
	}

	public void setListGT(List<GiamThi> listGT) {
		this.listGT = listGT;
	}

	public String getChucVu() {
		return chucVu;
	}

	public void setChucVu(String chucVu) {
		this.chucVu = chucVu;
	}

	public String getPhongThi() {
		return phongThi;
	}

	public void setPhongThi(String phongThi) {
		this.phongThi = phongThi;
	}

	public GiamThi() {
		// TODO Auto-generated constructor stub
	}

	public String getMaGV() {
		return maGV;
	}

	public void setMaGV(String maGV) {
		this.maGV = maGV;
	}

	public String getHoTen() {
		return hoTen;
	}

	public void setHoTen(String hoTen) {
		this.hoTen = hoTen;
	}

	public String getNgaySinh() {
		return ngaySinh;
	}

	public void setNgaySinh(String ngaySinh) {
		this.ngaySinh = ngaySinh;
	}

	public String getDonVi() {
		return donVi;
	}

	public void setDonVi(String donVi) {
		this.donVi = donVi;
	}
	
	
}
