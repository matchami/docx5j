package com.jumbletree.docx5j.xlsx.builders;

import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.xlsx4j.sml.Cell;
import org.xlsx4j.sml.STCellType;

import com.jumbletree.docx5j.xlsx.CommentPosition;

public class CellBuilder {

	private Cell cell;
	private RowBuilder parent;
	private WorkbookBuilder origin;
	private WorksheetBuilder sheet;
	private boolean unlocked;

	public CellBuilder(Cell cell, RowBuilder rowFactory, WorksheetBuilder sheet, WorkbookBuilder workbookBuilder) {
		this.cell = cell;
		this.parent = rowFactory;
		this.sheet = sheet;
		this.origin = workbookBuilder;
	}

	public CellBuilder style(String styleName) {
		cell.setS(unlocked ? origin.getUnlockedStyle(origin.getStyle(styleName)) : origin.getStyle(styleName));
		parent.checkThickBottom(styleName);
		return this;
	}

	/**
	 * Unlocks a cell in a protected sheet.  Will automatically switch on sheet protection if it isn't already
	 * @return
	 */
	public CellBuilder unlocked() throws Docx4JException {
		sheet.protect();
		this.unlocked = true;
		cell.setS(origin.getUnlockedStyle(cell.getS()));
		return this;
	}
	
	public CellBuilder value(String string) {
		cell.setT(STCellType.S);

		int index = origin.getStringCache(string);
		cell.setV(String.valueOf(index));

		return this;
	}

	public TextBuilder multiStyleText() {
		cell.setT(STCellType.S);
		int index = origin.createMultiStyleText();
		cell.setV(String.valueOf(index));
		
		return new TextBuilder(this, origin, cell.getS(), index);
	}
	
	public RowBuilder row() {
		return parent;
	}

	public CellBuilder value(double number) {
		cell.setV(String.valueOf(number));
		return this;
	}

	public CellBuilder comment(String author, String comment, CommentPosition position) throws Docx4JException {
		origin.createComment(sheet, cell, author, comment, position);
		
		return this;
	}

}
