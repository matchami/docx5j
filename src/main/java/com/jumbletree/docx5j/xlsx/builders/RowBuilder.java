package com.jumbletree.docx5j.xlsx.builders;

import java.util.List;

import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.xlsx4j.sml.CTMergeCell;
import org.xlsx4j.sml.CTMergeCells;
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
	
	public RowBuilder mergeCells(int start, int end) {
		CTMergeCells merges = getMergeCells();
		CTMergeCell merge = new CTMergeCell();
		XLSXRange range = XLSXRange.fromNumericReference(start, row.getR().intValue()-1, end, row.getR().intValue()-1);
		merge.setRef(range.rangeSheetlessReference());
		merges.getMergeCell().add(merge);
		merges.setCount((long) merges.getMergeCell().size());

		return this;
	}
	
	private CTMergeCells getMergeCells() {
		try {
			CTMergeCells cells = parent.sheet.getContents().getMergeCells();
			if (cells == null) { 
				cells = new CTMergeCells();
				parent.sheet.getContents().setMergeCells(cells);
			}
			return cells;
		} catch (Docx4JException e) {
			//Not realistically possible if we've got this far
			return null;
		}
	}

	public RowBuilder skipCell() {
		nextCell();
		return this;
	}
	
	public RowBuilder skipCells(int count) {
		for (int i=0; i<count; i++) {
			nextCell();
		}
		return this;
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
