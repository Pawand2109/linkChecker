package com.Pfizer.linkcheck;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.net.HttpURLConnection;
import java.net.URL;
import java.util.Iterator;
import java.util.List;
import java.util.concurrent.TimeUnit;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;

public class LinkChecker {

	public static WebDriver driver;
	public static String ioFile = "IOFiles//Pfizer Link Test Results.xls";
	public static Workbook wb;
	public static Sheet sheetObj;
	public static String urlToReport = "";

	public static void main(String[] args) throws Exception {
		appLaunch();
		readAndWriteToExcel();
		appQuit();

	}
	
	public static void appLaunch() {
		System.setProperty("webdriver.chrome.driver", "Drivers//chromedriver.exe");
		driver = new ChromeDriver();
		driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
		driver.manage().window().maximize();
	}

	public static String testBrokenlinks(String siteURL, Sheet sheetObj, int rownum) throws Exception {
		
		driver.get(siteURL);
		if (siteURL.contains("onmemoryca.test2")) {
			driver.switchTo().alert().dismiss();
		}

		List<WebElement> links = driver.findElements(By.tagName("a"));

		Iterator<WebElement> it = links.iterator();
		int counter = 0;
		int brokencnt = 1;
		urlToReport=null;
		urlToReport="";
		while (it.hasNext()) {
			try {
				counter++;
				String url = it.next().getAttribute("href");
				HttpURLConnection huc = (HttpURLConnection) (new URL(url).openConnection());
				huc.setRequestMethod("GET");
				huc.connect();
				int respCode = huc.getResponseCode();
				if (respCode == 404) {
					urlToReport = urlToReport + "\n" + brokencnt + "." + url;
					System.err
							.println("  " + counter + ". " + it.next().getText() + " --> " + url + " --> Broken Link");
					brokencnt++;
				} else {
					System.out.println(" " + counter + ". " + it.next().getText() + " --> " + url + " --> OK ");
				}
			} catch (Exception e) {
			}
		}
		return urlToReport;
	}

	

	public static void readAndWriteToExcel() throws Exception {
		FileInputStream fis = new FileInputStream(ioFile);

		wb = new HSSFWorkbook(fis);

		sheetObj = wb.getSheetAt(0);

		int rowlimit = sheetObj.getLastRowNum();

		for (int i = 1; i <= rowlimit; i++) {
			String testURL = null;
			Row row = sheetObj.getRow(i);
			Cell cell = row.getCell(0);
			testURL = cell.getStringCellValue().trim();
			if (!testURL.contains("http://") || !testURL.contains("https://")) {
				testURL = "http://" + testURL;
			}
			System.out.println(i + ". " + testURL);
			urlToReport=testBrokenlinks(testURL, sheetObj, i);
			writeToExcel(i,urlToReport);
		}

	}
	
	public static void writeToExcel(int rowtowrite,String brokenlinks) throws Exception {
          
		   Row rowObj = sheetObj.getRow(rowtowrite);
		   rowObj.createCell(1).setCellValue(brokenlinks);
		   OutputStream fos=new FileOutputStream(ioFile);
		   wb.write(fos);
	}
	
	
	public static void appQuit() {
		
		driver.quit();
		
	}
	
	

}
