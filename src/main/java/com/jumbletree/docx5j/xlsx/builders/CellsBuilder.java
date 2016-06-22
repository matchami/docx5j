package com.jumbletree.docx5j.xlsx.builders;

import java.util.List;

import org.xlsx4j.sml.Cell;

import com.jumbletree.docx5j.xlsx.XLSXRange;

/**
 * Builder for multiple cells in a row with the same attributes
 * @author matchami
 *
 */
public class CellsBuilder {

	private RowBuilder parent;
	private WorkbookBuilder origin;
	private int index;
	private String style;
	private int number;

	public CellsBuilder(int number, RowBuilder rowBuilder, WorkbookBuilder origin) {
		this.number = number;
		this.parent = rowBuilder;
		this.origin = origin;
		this.index = 0;
	}

	public CellsBuilder style(String styleName) {
		this.style = styleName;
		parent.checkThickBottom(styleName);
		return this;
	}

	public CellsBuilder values(String ... values) {
		for (int i=0; i<values.length; i++) {
			if (++index > number)
				throw new IndexOutOfBoundsException("More values provided than length of the cells");
			parent.nextCell().style(style).value(values[i]);
		}
		return this;
	}

	public CellsBuilder values(double ... values) {
		for (int i=0; i<values.length; i++) {
			if (++index > number)
				throw new IndexOutOfBoundsException("More values provided than length of the cells");
			parent.nextCell().style(style).value(values[i]);
		}
		return this;
	}

	public RowBuilder row() {
		for (int i=index; i<number; i++) {
			parent.nextCell().style(style).value("");
		}
		return parent;
	}
	
	

}
