package com.jumbletree.docx5j.xlsx.builders;

import java.awt.Color;

import org.xlsx4j.sml.CTCellAlignment;
import org.xlsx4j.sml.STBorderStyle;
import org.xlsx4j.sml.STHorizontalAlignment;
import org.xlsx4j.sml.STPatternType;
import org.xlsx4j.sml.STVerticalAlignment;

public class StyleBuilder {

	private Long formatId;
	private Long fontId;
	private Long fillId; 
	private Long borderId;
	private CTCellAlignment alignment;
	private WorkbookBuilder parent;
	private boolean thickBottom;
	
	public StyleBuilder(WorkbookBuilder parent) {
		this.parent = parent;
	}
	
	public StyleBuilder withFont(String fontName, int size, Color color, boolean bold, boolean italic) {
		this.fontId = parent.createFont(fontName, size, color, bold, italic);
		return this;
	}
	
	public StyleBuilder withFormat(int builtInFormat) {
		this.formatId = new Long(builtInFormat);
		return this;
	}
	
	public StyleBuilder withBorder(STBorderStyle style, Color color) {
		this.borderId = parent.createBorder(style, color);
		return this;
	}
	
	public StyleBuilder withBorder(STBorderStyle topStyle, Color topColor, STBorderStyle rightStyle, Color rightColor, STBorderStyle bottomStyle, Color bottomColor, STBorderStyle leftStyle, Color leftColor) {
		this.borderId = parent.createBorder(topStyle, topColor, rightStyle, rightColor, bottomStyle, bottomColor, leftStyle, leftColor);
		this.thickBottom = false;
		if (bottomStyle != null) {
			//See https://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.row(v=office.14).aspx
			switch (bottomStyle) {
				case MEDIUM_DASH_DOT_DOT:
				case SLANT_DASH_DOT:
				case MEDIUM_DASH_DOT:
				case MEDIUM_DASHED:
				case MEDIUM:
				case THICK:
				case DOUBLE:
					this.thickBottom = true;
					break;
				default:
			}
		}
		return this;
	}

	public boolean hasThickBottom() {
		return thickBottom;
	}

	public void setThickBottom(boolean thickBottom) {
		this.thickBottom = thickBottom;
	}

	public StyleBuilder withFill(Color bgColor, Color fgColor, STPatternType pattern) {
		this.fillId = parent.createFill(bgColor, fgColor, pattern);
		return this;
	}
	
	public StyleBuilder withAlignment(STHorizontalAlignment horizontal, STVerticalAlignment vertical) {
		CTCellAlignment alignment = new CTCellAlignment();
		alignment.setHorizontal(horizontal);
		alignment.setVertical(vertical);
		this.alignment = alignment;
		
		return this;
	}

	public StyleBuilder installAs(String name) {
		int index = parent.createStyle(formatId, fontId, fillId, borderId, alignment);
		parent.installStyle(name, index);
		if (thickBottom)
			parent.installThickBottomStyle(index);
		return this;
	}
	
	/**
	 * Doesn't actually copy anything.  This should only be called after installAs to make a second style with similar attributes
	 * @return
	 */
	public StyleBuilder copy() {
		return this;
	}
}
