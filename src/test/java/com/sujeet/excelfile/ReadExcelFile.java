package com.sujeet.excelfile;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;



public class ReadExcelFile {
	public static void main(String[] args) throws IOException, InterruptedException {
		
		
		
		WebDriver driver=new ChromeDriver();
		driver.manage().window().maximize();
		driver.get("https://admin-demo.nopcommerce.com/login");
		Thread.sleep(2000);
		
		String filepath="C:\\Sujeet Yadav\\DataDrivenTesting.xlsx";
		FileInputStream file=new FileInputStream(filepath);
		XSSFWorkbook workbook=new XSSFWorkbook(file);
		XSSFSheet sheet=workbook.getSheet("Data");
		int rows=sheet.getLastRowNum();
		System.out.println(rows);
		for(int r=1;r<=rows;r++) {
			XSSFRow row=sheet.getRow(r);
			XSSFCell emailId=row.getCell(0);
			XSSFCell password=row.getCell(1);
			System.err.println("Email as-->"+emailId+" password-->"+password);
			
			try {
					
				driver.findElement(By.id("Email")).clear();
				driver.findElement(By.id("Email")).sendKeys(emailId.toString());
				driver.findElement(By.id("Password")).clear();
				driver.findElement(By.id("Password")).sendKeys(password.toString());
				driver.findElement(By.xpath("//button[@type='submit']")).click();
				Thread.sleep(2000);
				
				driver.findElement(By.linkText("Logout")).click();
				Thread.sleep(2000);
				System.out.println("valid");
				Thread.sleep(2000);
				
			}
			catch(Exception e) {
				System.out.println("invalid");
			}
			}
			
		}

}
