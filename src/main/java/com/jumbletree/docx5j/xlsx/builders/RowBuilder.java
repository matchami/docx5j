package com.jumbletree.docx5j.xlsx.builders;

import java.util.List;

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

	public RowBuilder height(double height, boolean ... isCustom) {
		row.setHt(height);
		if (isCustom.length > 0)
			row.setCustomHeight(isCustom[0]);
		
		return this;
	}

	public CellBuilder nextCell() {
		List<Cell> cells = row.getC();
		XLSXRange cellRange = XLSXRange.fromNumericReference(cells.size(), row.getR().intValue()-1);
		
		Cell cell = new Cell();
		cell.setR(cellRange.singleCellSheetlessReference());
		
		cells.add(cell);
		return new CellBuilder(cell, this, parent, origin);
	}

	public RowBuilder addExplicitSpan(int min, int max) {
		row.getSpans().add(min + ":" + max);
		return this;
	}

	public WorksheetBuilder sheet() {
		return parent;
	}

	public RowBuilder style(String styleName) {
		row.setS(origin.getStyle(styleName));
		checkThickBottom(styleName);
		return this;
	}

	void checkThickBottom(String styleName) {
		if (origin.isThickBottomStyle(origin.getStyle(styleName))) {
			row.setThickBot(true);
		}
	}

	public CellsBuilder cells(int number) {
		return new CellsBuilder(number, this, origin);
	}

	public WorksheetBuilder repeat(int repeats) {

		return parent;
	}



}
