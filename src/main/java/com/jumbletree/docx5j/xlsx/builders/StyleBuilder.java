package com.jumbletree.docx5j.xlsx.builders;

import java.awt.Color;

import com.jumbletree.docx5j.xlsx.XLSXFile;

public class StyleBuilder {

	Long formatId;
	Long fontId;
	Long fillId; 
	Long borderId;
	private WorkbookBuilder parent;
	
	public StyleBuilder(WorkbookBuilder parent) {
		this.parent = parent;
	}
	
	public StyleBuilder withFont(String fontName, int size, Color color, boolean bold, boolean italic) {
		this.fontId = new Long(parent.createFont(fontName, size, color, bold, italic));
		return this;
	}
	
	public StyleBuilder withFormat() {
		//TODO
		return this;
	}
	
	public StyleBuilder withBorder() {
		//TODO
		return this;
	}

	public StyleBuilder withFill() {
		//TODO
		return this;
	}

	public void installAs(String name) {
		int index = parent.createStyle(formatId, fontId, fillId, borderId);
		parent.installStyle(name, index);
	}
}
