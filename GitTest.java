package com.example.tests;

import java.io.FileInputStream;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;

public class GitTest {
	public static void main(String[] args) throws Exception{
		//读入xslx文件
		InputStream is = new FileInputStream("C:\\Users\\admin\\Downloads\\input.xlsx");
        XSSFWorkbook xssfWorkbook = new XSSFWorkbook(is);
        XSSFSheet sheet = xssfWorkbook.getSheetAt(0);
        Row r = sheet.getRow(0);
        String num = null;
        String path = null;
        //存储数据
        List<Map<String,String>> list = new ArrayList<Map<String,String>>();
        Map<String, String> content = new HashMap<String, String>();
        for(int i = 0; i < 96; i++) {
            r = sheet.getRow(i);
            if(r != null){
            	num = (String) getCellFormatValue(r.getCell(0));
                path = (String) getCellFormatValue(r.getCell(1));
                content.put(num, path);
                
            }else{
                break;
            }
            list.add(content);
        }
        
		System.getProperty("webdriver.chrome.driver", "C:\\Program Files (x86)\\Google\\Chrome\\Application\\chromedriver.exe");
		WebDriver driver = new ChromeDriver();
		driver.get("https://psych.liebes.top/st");
		
		for(Entry<String, String> entry : content.entrySet()){
			String stuNum = entry.getKey();
        	String stuGit = entry.getValue();
        	driver.findElement(By.id("username")).click();
            driver.findElement(By.id("username")).clear();
            driver.findElement(By.id("username")).sendKeys(stuNum);
            driver.findElement(By.id("password")).click();
            driver.findElement(By.id("password")).clear();
            driver.findElement(By.id("password")).sendKeys(stuNum.substring(4,10));
            driver.findElement(By.id("submitButton")).click();
            if(stuGit.equals(driver.findElement(By.xpath("//p")).getText())){
            	System.out.println("学号与git匹配");
            }else{
            	System.out.println("学号与git不匹配");
            }
		}
		
		driver.close();
	}
	
	private static Object getCellFormatValue(Cell cell){  
        Object cellvalue = "";  
        if (cell != null) {  
            // 判断当前Cell的Type  
            switch (cell.getCellType()) {  
            case Cell.CELL_TYPE_NUMERIC:// 如果当前Cell的Type为NUMERIC  
            case Cell.CELL_TYPE_FORMULA: {  
                // 判断当前的cell是否为Date  
                if (DateUtil.isCellDateFormatted(cell)) {  
                    // 如果是Date类型则，转化为Data格式  
                    // data格式是带时分秒的：2013-7-10 0:00:00  
                    // cellvalue = cell.getDateCellValue().toLocaleString();  
                    // data格式是不带带时分秒的：2013-7-10  
                    Date date = cell.getDateCellValue();  
                    cellvalue = date;  
                } else {// 如果是纯数字  
  
                    // 取得当前Cell的数值  
                    cellvalue = String.valueOf(cell.getNumericCellValue());  
                }  
                break;  
            }  
            case Cell.CELL_TYPE_STRING:// 如果当前Cell的Type为STRING  
                // 取得当前的Cell字符串  
                cellvalue = cell.getRichStringCellValue().getString();  
                break;  
            default:// 默认的Cell值  
                cellvalue = "";  
            }  
        } else {  
            cellvalue = "";  
        }  
        return cellvalue;  
    }  
}
