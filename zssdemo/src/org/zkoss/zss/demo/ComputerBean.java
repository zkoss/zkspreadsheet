package org.zkoss.zss.demo;

import java.io.Serializable;

public class ComputerBean implements Serializable {
	private String id = "";
	private String product = "";
	private String brand = "";
	private String model = "";
	private String serialNumber = "";
	private String date = "";
	private String warrantyTime = "";
	private double cost = 0.0;
	private String os = "";
	private double salvage = 0.0;
	public String getId() {
		return id;
	}
	public void setId(String id) {
		this.id = id;
	}
	public String getProduct() {
		return product;
	}
	public void setProduct(String product) {
		this.product = product;
	}
	public String getBrand() {
		return brand;
	}
	public void setBrand(String brand) {
		this.brand = brand;
	}
	public String getModel() {
		return model;
	}
	public void setModel(String model) {
		this.model = model;
	}
	public String getSerialNumber() {
		return serialNumber;
	}
	public void setSerialNumber(String serialNumber) {
		this.serialNumber = serialNumber;
	}
	public String getDate() {
		return date;
	}
	public void setDate(String date) {
		this.date = date;
	}
	public String getWarrantyTime() {
		return warrantyTime;
	}
	public void setWarrantyTime(String warrantyTime) {
		this.warrantyTime = warrantyTime;
	}
	public double getCost() {
		return cost;
	}
	public void setCost(double cost) {
		this.cost = cost;
	}
	public String getOs() {
		return os;
	}
	public void setOs(String os) {
		this.os = os;
	}
	public double getSalvage() {
		return salvage;
	}
	public void setSalvage(double salvage) {
		this.salvage = salvage;
	}
	
	
}
