package com.jumbletree.docx5j.xlsx.builders;

import java.awt.Color;

import org.xlsx4j.sml.CTCellAlignment;
import org.xlsx4j.sml.STBorderStyle;
import org.xlsx4j.sml.STHorizontalAlignment;
import org.xlsx4j.sml.STPatternType;
import org.xlsx4j.sml.STUnderlineValues;
import org.xlsx4j.sml.STVerticalAlignment;

public class StyleBuilder {

	public static enum Edge {
		TOP(1), TOP_RIGHT(3), RIGHT(2), BOTTOM_RIGHT(6), BOTTOM(4), BOTTOM_LEFT(12), LEFT(8), TOP_LEFT(9),
		TOP_BOTTOM(5), LEFT_RIGHT(10),
		NOT_TOP(14), NOT_RIGHT(13), NOT_BOTTOM(11), NOT_LEFT(7),
		ALL(15);
		
		int sides;
		
		Edge(int sides) {
			this.sides = sides;
		}
		
		public boolean hasLeft() {
			return (sides & LEFT.sides) > 0;
		}
		
		public boolean hasRight() {
			return (sides & RIGHT.sides) > 0;
		}
		
		public boolean hasTop() {
			return (sides & TOP.sides) > 0;
		}
		
		public boolean hasBottom() {
			return (sides & BOTTOM.sides) > 0;
		}
	}
	
	private Long formatId;
	private Long fontId;
	private Long fillId; 
	private Long borderId;
	private CTCellAlignment alignment;
	private WorkbookBuilder parent;
	private boolean thickBottom;
	private String formatDefinition;
	
	public StyleBuilder(WorkbookBuilder parent) {
		this.parent = parent;
	}
	
	public StyleBuilder withFont(String fontName, int size, Color color) {
		return withFont(fontName, size, color, false, false, null);
	}
	
	public StyleBuilder withFont(String fontName, int size, Color color, boolean bold, boolean italic) {
		return withFont(fontName, size, color, bold, italic, null);
	}
	
	public StyleBuilder withFont(String fontName, int size, Color color, boolean bold, boolean italic, STUnderlineValues underline) {
		this.fontId = parent.createFont(fontName, size, color, bold, italic, underline);
		return this;
	}
	
	public StyleBuilder withFormat(int builtInFormat) {
		this.formatId = Long.valueOf(builtInFormat);
		return this;
	}
	
	public StyleBuilder withFormat(String formatDefinition) {
		this.formatId = null;
		this.formatDefinition = formatDefinition;
		return this;
	}

	public StyleBuilder withBorder(STBorderStyle style, Color color) {
		this.borderId = parent.createBorder(style, color);
		return this;
	}
	
	public StyleBuilder withBorder(STBorderStyle style, Edge edge) {
		return withBorder(style, null, edge);
	}

	public StyleBuilder withBorder(STBorderStyle style, Color color, Edge edge) {
		STBorderStyle topStyle = null; 
		Color topColor = null; 
		STBorderStyle rightStyle = null; 
		Color rightColor = null; 
		STBorderStyle bottomStyle = null; 
		Color bottomColor = null; 
		STBorderStyle leftStyle = null; 
		Color leftColor = null;

		if (edge.hasTop()) {
			topStyle = style; 
			topColor = color;
		}
		if (edge.hasRight()) {
			rightStyle = style; 
			rightColor = color;
		}
		if (edge.hasBottom()) {
			bottomStyle = style; 
			bottomColor = color;
		}
		if (edge.hasLeft()) {
			leftStyle = style; 
			leftColor = color;
		}
		
		return withBorder(topStyle, topColor, rightStyle, rightColor, bottomStyle, bottomColor, leftStyle, leftColor);
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

	public StyleBuilder clearBorders() {
		this.borderId = null;
		this.thickBottom = false;
		return this;
	}
	
	public boolean hasThickBottom() {
		return thickBottom;
	}

	public void setThickBottom(boolean thickBottom) {
		this.thickBottom = thickBottom;
	}

	public StyleBuilder withFill(Color color) {
		this.fillId = parent.createFill(color, color, STPatternType.SOLID);
		return this;
	}
	
	public StyleBuilder clearFill() {
		this.fillId = null;
		return this;
	}

	public StyleBuilder withFill(Color bgColor, Color fgColor, STPatternType pattern) {
		this.fillId = parent.createFill(bgColor, fgColor, pattern);
		return this;
	}
	
	public StyleBuilder withAlignment(STHorizontalAlignment horizontal, STVerticalAlignment vertical) {
		withAlignment(horizontal, vertical, false);
		return this;
	}

	public StyleBuilder withAlignment(STHorizontalAlignment horizontal, STVerticalAlignment vertical, boolean wrapText) {
		CTCellAlignment alignment = new CTCellAlignment();
		alignment.setHorizontal(horizontal);
		alignment.setVertical(vertical);
		this.alignment = alignment;
		this.alignment.setWrapText(wrapText);
		return this;
	}
	
    public StyleBuilder withAlignment(STHorizontalAlignment horizontal, STVerticalAlignment vertical, boolean wrapText,
            long textRotation) {
        CTCellAlignment alignment = new CTCellAlignment();
        alignment.setHorizontal(horizontal);
        alignment.setVertical(vertical);
        alignment.setTextRotation(textRotation);
        this.alignment = alignment;
        this.alignment.setWrapText(wrapText);
        return this;
    }

	
	
	public StyleBuilder installAs(String name) {
		int index = formatId == null && formatDefinition != null ?
				parent.createStyle(formatDefinition, fontId, fillId, borderId, alignment) :
				parent.createStyle(formatId, fontId, fillId, borderId, alignment);
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
