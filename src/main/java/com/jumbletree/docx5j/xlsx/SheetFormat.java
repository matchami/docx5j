package com.jumbletree.docx5j.xlsx;

import org.xlsx4j.sml.CTSheetFormatPr;

public class SheetFormat {

	private double defaultRowHeight;

	public SheetFormat(double defaultRowHeight) {
		this.defaultRowHeight = defaultRowHeight;
	}

	public CTSheetFormatPr createCTSheetFormatPr() {
		CTSheetFormatPr pr = new CTSheetFormatPr();
		pr.setDefaultRowHeight(defaultRowHeight);
		return pr;
	}

}
