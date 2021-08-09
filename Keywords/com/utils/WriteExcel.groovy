package com.utils

import static com.kms.katalon.core.checkpoint.CheckpointFactory.findCheckpoint
import static com.kms.katalon.core.testcase.TestCaseFactory.findTestCase
import static com.kms.katalon.core.testdata.TestDataFactory.findTestData
import static com.kms.katalon.core.testobject.ObjectRepository.findTestObject
import static com.kms.katalon.core.testobject.ObjectRepository.findWindowsObject

import org.apache.poi.xssf.usermodel.XSSFCell
import org.apache.poi.xssf.usermodel.XSSFRow
import org.apache.poi.xssf.usermodel.XSSFSheet
import org.apache.poi.xssf.usermodel.XSSFWorkbook

import com.kms.katalon.core.annotation.Keyword
import com.kms.katalon.core.checkpoint.Checkpoint
import com.kms.katalon.core.configuration.RunConfiguration
import com.kms.katalon.core.cucumber.keyword.CucumberBuiltinKeywords as CucumberKW
import com.kms.katalon.core.mobile.keyword.MobileBuiltInKeywords as Mobile
import com.kms.katalon.core.model.FailureHandling
import com.kms.katalon.core.testcase.TestCase
import com.kms.katalon.core.testdata.TestData
import com.kms.katalon.core.testobject.TestObject
import com.kms.katalon.core.webservice.keyword.WSBuiltInKeywords as WS
import com.kms.katalon.core.webui.keyword.WebUiBuiltInKeywords as WebUI
import com.kms.katalon.core.windows.keyword.WindowsBuiltinKeywords as Windows

import internal.GlobalVariable

public class WriteExcel {

	@Keyword
	public void demoWriteExcel(String strTest) throws IOException{
		String path = RunConfiguration.getProjectDir();
		FileInputStream fis = new FileInputStream(path + "/TestDetails/Excel_Automation.xlsx")
		XSSFWorkbook wb = new XSSFWorkbook(fis);
		XSSFSheet sheet = wb.getSheet('Sheet1');
		XSSFRow row = sheet.getRow(1);
		int colNum = row.getLastCellNum();
		print('Total number of columns :'+colNum);
		print('Value in strTest is :'+strTest);
		XSSFCell cell = null;
		if (cell == null)
			cell = row.createCell(colNum);
		cell.setCellValue(strTest);

		FileOutputStream fos = new FileOutputStream(path + "/TestDetails/Excel_Automation.xlsx");
		wb.write(fos);
		fos.close();
	}
}
