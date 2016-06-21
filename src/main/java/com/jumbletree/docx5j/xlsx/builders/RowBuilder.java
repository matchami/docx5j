package com.jumbletree.docx5j.xlsx.builders;

import java.util.List;

import org.xlsx4j.sml.CTExtension;
import org.xlsx4j.sml.CTExtensionList;
import org.xlsx4j.sml.Cell;
import org.xlsx4j.sml.Row;

import com.jumbletree.docx5j.xlsx.XLSXRange;

public class RowBuilder {

	private Row row;
	private WorksheetBuilder parent;
	private WorkbookBuilder origin;

	public RowBuilder(Row row, WorksheetBuilder worksheetFactory, WorkbookBuilder workbookBuilder) {
		this.row = row;
		this.parent = worksheetFactory;
		this.origin = workbookBuilder;
	}

	public RowBuilder setHeight(double height) {
		row.setHt(height);
		return this;
	}

	public CellBuilder nextCell() {
		List<Cell> cells = row.getC();
		XLSXRange cellRange = XLSXRange.fromNumericReference(cells.size()+1, row.getR().intValue()-1);
		
		Cell cell = new Cell();
		cell.setR(cellRange.singleCellSheetlessReference());
		
		cells.add(cell);
		return new CellBuilder(cell, this, origin);
	}

	public RowBuilder addExplicitSpan(int min, int max) {
		row.getSpans().add(min + ":" + max);
		return this;
	}

	public WorksheetBuilder sheet() {
		return parent;
	}

}
