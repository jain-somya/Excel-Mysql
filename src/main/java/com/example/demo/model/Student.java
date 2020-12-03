package com.example.demo.model;

import javax.persistence.Entity;
import javax.persistence.Id;

@Entity
public class Student {
@Id
private int sId;
private String sName;
private String sEmail;
public int getsId() {
	return sId;
}
public void setsId(int sId) {
	this.sId = sId;
}
public String getsName() {
	return sName;
}
public void setsName(String sName) {
	this.sName = sName;
}
public String getsEmail() {
	return sEmail;
}
public void setsEmail(String sEmail) {
	this.sEmail = sEmail;
}

}
