package com.jumbletree.docx5j.xlsx;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.HashMap;

import javax.xml.bind.JAXBException;

import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.exceptions.InvalidFormatException;

import com.jumbletree.docx5j.xlsx.builders.StyleBuilder;
import com.jumbletree.docx5j.xlsx.builders.WorkbookBuilder;

public class XLSXFile {

	private WorkbookBuilder factory;
	
	public XLSXFile() throws InvalidFormatException, JAXBException {
		this.factory = new WorkbookBuilder();
	}
	
	public WorkbookBuilder getWorkbookBuilder() {
		return factory;
	}
	
	public void save(File toFile) throws IOException, Docx4JException {
		FileOutputStream out = new FileOutputStream(toFile);
		save(out);
		out.close();
	}

	public void save(OutputStream out) throws IOException, Docx4JException {
		factory.save(out);
	}
	
	public StyleBuilder createStyle() {
		return new StyleBuilder(factory);
	}
}
