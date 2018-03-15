package com.jumbletree.docx5j.xlsx.builders;

import java.awt.Color;
import java.io.InputStream;
import java.util.List;

import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.parts.SpreadsheetML.WorksheetPart;
import org.xlsx4j.sml.CTBreak;
import org.xlsx4j.sml.CTColor;
import org.xlsx4j.sml.CTMergeCell;
import org.xlsx4j.sml.CTMergeCells;
import org.xlsx4j.sml.CTPageBreak;
import org.xlsx4j.sml.CTPageMargins;
import org.xlsx4j.sml.CTSheetDimension;
import org.xlsx4j.sml.CTSheetPr;
import org.xlsx4j.sml.Col;
import org.xlsx4j.sml.Cols;
import org.xlsx4j.sml.Row;
import org.xlsx4j.sml.STOrientation;
import org.xlsx4j.sml.SheetViews;
import org.xlsx4j.sml.Worksheet;

import com.jumbletree.docx5j.xlsx.SheetFormat;
import com.jumbletree.docx5j.xlsx.View;
import com.jumbletree.docx5j.xlsx.XLSXRange;

public class WorksheetBuilder implements BuilderMethods {

	WorksheetPart sheet;
	private WorkbookBuilder parent;
	private int index;

	public WorksheetBuilder(int index, WorksheetPart worksheetPart, WorkbookBuilder parent) {
		this.index = index;
		this.sheet = worksheetPart;
		this.parent = parent;
	}

	public WorksheetBuilder setTabColor(Color color) throws Docx4JException {
		Worksheet worksheet = getWorksheet();
		CTSheetPr pr = worksheet.getSheetPr();
		if (pr == null) {
			pr = new CTSheetPr();
			worksheet.setSheetPr(pr);
		}
		CTColor ctcolor = new CTColor();
		ctcolor.setRgb(getColorBytes(color));
		pr.setTabColor(ctcolor);
		return this;
	}

	public Worksheet getWorksheet() throws Docx4JException {
		return sheet.getContents();
	}

	public WorksheetBuilder setName(String name) throws Docx4JException {
		parent.setSheetName(index, name);
		return this;
	}

	public WorksheetBuilder addView(View view) throws Docx4JException {
		Worksheet worksheet = getWorksheet();
		SheetViews views = worksheet.getSheetViews();
		if (views == null) {
			views = new SheetViews();
			worksheet.setSheetViews(views);
		}
		
		views.getSheetView().add(view.createSheetView());
		return this;
	}

	public WorksheetBuilder setDimension(XLSXRange range) throws Docx4JException {
		CTSheetDimension dim = new CTSheetDimension();
		dim.setRef(range.rangeSheetlessReference());
		getWorksheet().setDimension(dim);
		return this;
	}

	public WorksheetBuilder setFormat(SheetFormat sheetFormat) throws Docx4JException {
		getWorksheet().setSheetFormatPr(sheetFormat.createCTSheetFormatPr());
		return this;
	}

	/**
	 * 
	 * @param width
	 * @param bestFit
	 * @param columnRange if none, the next column will be used, if a single value, the give
	 * 			value will be used for the min and the max, if two values, will be used as
	 * 			min and max in that order.
	 * @return
	 * @throws Docx4JException 
	 */
	public WorksheetBuilder addColumnDefinition(double width, boolean bestFit, long ... columnRange) throws Docx4JException {
		Col col = new Col();
		col.setCustomWidth(true);
		col.setWidth(width);
		col.setBestFit(bestFit);
		
		Long overrideMin = null, overrideMax = null;
		
		if (columnRange.length > 0) {
			overrideMin = columnRange[0];
			if (columnRange.length > 1) {
				overrideMax = columnRange[1];
			}
		}
		
		List<Cols> colsList = getWorksheet().getCols();
		Cols cols = null;
		if (colsList.size() == 0) {
			cols = new Cols();
			colsList.add(cols);
		} else {
			cols = colsList.get(0);
		}
		if (overrideMin == null) {
			long min = 0;
			for (Col aCol : cols.getCol()) {
				min = Math.max(min, aCol.getMax());
			}
			overrideMin = min + 1;
		}
		if (overrideMax == null) {
			overrideMax = overrideMin;
		}
		
		col.setMin(overrideMin);
		col.setMax(overrideMax);
		cols.getCol().add(col);
		
		return this;
	}

	public RowBuilder nextRow() throws Docx4JException {
		List<Row> rows = sheet.getContents().getSheetData().getRow();
		
		Row row = new Row();
		row.setR(new Long(rows.size()+1));
		rows.add(row);
		
		return new RowBuilder(row, this, parent);
	}

	public WorksheetBuilder setPageMargns(double header, double footer, double top, double right, double bottom, double left) throws Docx4JException {
		CTPageMargins margins = new CTPageMargins();
		margins.setHeader(header);
		margins.setBottom(bottom);
		margins.setTop(top);
		margins.setLeft(left);
		margins.setRight(right);
		margins.setFooter(footer);
		sheet.getContents().setPageMargins(margins);
		return this;
	}
	
	public WorksheetBuilder setPageSetup(InputStream data, STOrientation orientation, int scale, int paperSize) throws Docx4JException {
		parent.createPageSetup(this, data, orientation, scale, paperSize);
		return this;
	}

	public WorksheetBuilder addRowBreak(int after) throws Docx4JException {
		return addRowBreak(after, null);
	}

	public WorksheetBuilder addRowBreak(int after, Integer max) throws Docx4JException {
		Worksheet sheet = this.sheet.getContents();
		CTPageBreak breaks = sheet.getRowBreaks();
		if (breaks == null) {
			sheet.setRowBreaks(breaks = new CTPageBreak());
			breaks.setCount(new Long(0));
			breaks.setManualBreakCount(new Long(0));
		}
		breaks.setCount(new Long(breaks.getCount() + 1));
		breaks.setManualBreakCount(new Long(breaks.getManualBreakCount() + 1));
		CTBreak brk = new CTBreak();
		brk.setMan(true);
		brk.setId(new Long(after));
		if (max != null)
			brk.setMax(max.longValue());
		breaks.getBrk().add(brk);

		return this;
	}
	
	public WorksheetBuilder addMergeCell(String merge) throws Docx4JException {
		Worksheet worksheet = getWorksheet();
		CTMergeCell mergeC = new CTMergeCell();
		mergeC.setRef(merge);
		if (worksheet.getMergeCells() == null) {
			worksheet.setMergeCells(new CTMergeCells());
		}
		worksheet.getMergeCells().getMergeCell().add(mergeC);
		worksheet.getMergeCells().setCount((long) worksheet.getMergeCells().getMergeCell().size());
		return this;
	}
}
