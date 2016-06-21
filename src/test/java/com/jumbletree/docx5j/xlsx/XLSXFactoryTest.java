package com.jumbletree.docx5j.xlsx;

import static org.junit.Assert.assertEquals;

import javax.xml.bind.JAXBException;

import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.junit.Test;

import com.jumbletree.docx5j.xlsx.XLSXFile;

public class XLSXFactoryTest {

	@Test
	public void testAppendSheet() throws Docx4JException, JAXBException {
		XLSXFile file = new XLSXFile();
		
		//Sheet1
		assertEquals("Default sheet was not called \"Sheet1\"", "Sheet1", file.getWorkbookBuilder().getSheetName(0));
		file.getWorkbookBuilder().appendSheet();
		//Sheet1,Sheet2
		assertEquals("Appended sheet was not called \"Sheet2\"", "Sheet2", file.getWorkbookBuilder().getSheetName(1));
		file.getWorkbookBuilder().appendSheet();
		//Sheet1,Sheet2,Sheet3
		assertEquals("Appended sheet was not called \"Sheet3\"", "Sheet3", file.getWorkbookBuilder().getSheetName(2));
		//Sheet1,Bob,Sheet3
		file.getWorkbookBuilder().getSheet(1).setName("Bob");
		file.getWorkbookBuilder().appendSheet();
		//Sheet1,Bob,Sheet3,Sheet4
		assertEquals("Appended sheet was not called \"Sheet4\"", "Sheet4", file.getWorkbookBuilder().getSheetName(3));
		file.getWorkbookBuilder().getSheet(3).setName("George");
		file.getWorkbookBuilder().getSheet(2).setName("Mildred");
		file.getWorkbookBuilder().appendSheet();
		//Sheet1,Bob,Mildred,George,Sheet2
		assertEquals("Appended sheet was not called \"Sheet2\"", "Sheet2", file.getWorkbookBuilder().getSheetName(4));
		file.getWorkbookBuilder().getSheet(0).setName("Margo");
		file.getWorkbookBuilder().getSheet(4).setName("Victor");
		file.getWorkbookBuilder().appendSheet();
		//Margo,Bob,Mildred,George,Victor,Sheet1
		assertEquals("Appended sheet was not called \"Sheet1\"", "Sheet1", file.getWorkbookBuilder().getSheetName(5));
		
	}
}
