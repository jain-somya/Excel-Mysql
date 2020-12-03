package com.example.demo.service;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.List;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import com.example.demo.repos.StudentRepo;
import com.example.demo.model.Student;

@Service
public class ExcelDb {
	@Autowired
	StudentRepo repo;
	public  List<Student> FromExceltoDb(InputStream is) throws EncryptedDocumentException, IOException{
		try{
			Workbook workbook = new XSSFWorkbook(is);
		 for(Sheet sheet: workbook) {
			 for (Row row: sheet) {
				 String email = row.getCell(2).getStringCellValue();
	            	int id =  (int) row.getCell(0).getNumericCellValue();
	            	String sname =  row.getCell(1).getStringCellValue();
	            	Student st =new Student();
	            	st.setsId(id);
	            	st.setsName(sname);
	            	st.setsEmail(email);
	            	repo.save(st);
			 }
		 }
		 workbook.close();
		}
		catch(IOException e){
			
		}
		return null;
	}
	
	public  ByteArrayInputStream FromDbtoExcel() throws IOException{
		
		ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
		List<Student> listUsers = repo.findAll();
		System.out.println(listUsers);
		XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("Students");
        int rowCount=0;
		for (Student st : listUsers) {
			Row row = sheet.createRow(rowCount++);
			int sId=st.getsId();
			String sEmail=st.getsEmail();
			String sName=st.getsName();
			Cell cell = row.createCell(0);
            cell.setCellValue(sId);
            
            cell = row.createCell(1);
            cell.setCellValue(sName);
            
            cell = row.createCell(2);
            cell.setCellValue(sEmail);
            
		}
		workbook.write(outputStream);
		workbook.close();
		return new ByteArrayInputStream(outputStream.toByteArray());
	}
}
