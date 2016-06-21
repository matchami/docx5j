package com.jumbletree.docx5j.xlsx.builders;

import org.xlsx4j.sml.CTRst;
import org.xlsx4j.sml.CTXstringWhitespace;
import org.xlsx4j.sml.Cell;
import org.xlsx4j.sml.STCellType;

public class CellBuilder {

	private Cell cell;
	private RowBuilder parent;
	private WorkbookBuilder origin;

	public CellBuilder(Cell cell, RowBuilder rowFactory, WorkbookBuilder workbookBuilder) {
		this.cell = cell;
		this.parent = rowFactory;
		this.origin = workbookBuilder;
	}

	public CellBuilder style(String styleName) {
		cell.setS(origin.getStyle(styleName));
		return this;
	}

	public CellBuilder value(String string) {
		cell.setT(STCellType.S);

		int index = origin.getStringCache(string);
		cell.setV(String.valueOf(index));

		return this;
	}

	public RowBuilder row() {
		return parent;
	}

}
