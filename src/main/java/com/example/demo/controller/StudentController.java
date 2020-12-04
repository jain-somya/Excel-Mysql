package com.example.demo.controller;

import java.io.ByteArrayInputStream;
import java.io.IOException;

import org.springframework.http.HttpHeaders;

import javax.servlet.http.HttpServletResponse;
import org.springframework.http.MediaType;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.core.io.InputStreamResource;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.multipart.MultipartFile;

import com.example.demo.service.ExcelDb;

@Controller
public class StudentController {
	
	@Autowired
	private ExcelDb service;
	
	@GetMapping("/")
	public String fp() {
		
		 
		 return "index.html";
	}
	@PostMapping("/toDb")
	public String toDb(@RequestParam("file") MultipartFile file) {
		 try {service.FromExceltoDb(file.getInputStream());}
		 catch(Exception e) {}
		 
		 return "home.html";
	}
	
	@GetMapping("/toExcel")
	public ResponseEntity toExcel(HttpServletResponse response) throws IOException {
		ByteArrayInputStream in;
		try {
			in = service.FromDbtoExcel();
			
		}
		 catch(Exception e) {in = service.FromDbtoExcel();}
		HttpHeaders headers = new HttpHeaders();
		headers.add("Content-Disposition", "attachment; filename=Students.xlsx");
		 
		return ResponseEntity.ok().headers(headers)
				.contentType(MediaType.parseMediaType("application/vnd.ms-excel"))
				.body(new InputStreamResource(in));
	}
}
