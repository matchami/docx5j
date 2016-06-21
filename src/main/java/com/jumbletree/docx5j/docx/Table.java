package com.jumbletree.docx5j.docx;

import java.awt.Insets;
import java.awt.Point;
import java.math.BigInteger;
import java.util.Arrays;
import java.util.HashMap;

import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.wml.CTHeight;
import org.docx4j.wml.CTShd;
import org.docx4j.wml.STHeightRule;
import org.docx4j.wml.Tbl;
import org.docx4j.wml.TblBorders;
import org.docx4j.wml.TblGrid;
import org.docx4j.wml.TblGridCol;
import org.docx4j.wml.TblPr;
import org.docx4j.wml.TblWidth;
import org.docx4j.wml.Tc;
import org.docx4j.wml.TcPr;
import org.docx4j.wml.TcPrInner.GridSpan;
import org.docx4j.wml.TcPrInner.TcBorders;
import org.docx4j.wml.Tr;
import org.docx4j.wml.TrPr;

public class Table {

	private DOCXBuilder factory;

	//Defaults...
	private int indent = 0;
	private Insets cellMargin = new Insets(0, 108, 0, 108);
	
	//Need to be set explicitly...
	private int[] columnWidths;
	private String[] cellFills;
//	private String[] cellStyles;
	
	/**
	 * Each cell's value.  May contain HTML b and i tags and may contain \n in which case each
	 * line is treated as a separate paragraph.  in this case HTML tags must resolve within each line.
	 */
	private String[][] data;

	//Set by constructor
	private int cols;
	private int rows;

	//Added to as necessary
	private HashMap<Point, Integer> hMerges = new HashMap<Point, Integer>();

	private String[] cellStyles;

	private Integer[] rowHeights;

	private String[][] borders;

	private String borderCSS;
	
	public Table(DOCXBuilder factory, String[][] data) {
		this.factory = factory;
		this.cols = data.length;
		this.rows = data[0].length;
		this.data = data;
	}
	
	public Table setIndent(int indent) {
		this.indent = indent;
		
		return this;
	}
	
	public Table setCellMargins(Insets margins) {
		this.cellMargin = margins;
		
		return this;
	}

	public Table setColumnWidths(int[] widths) {
		if (widths.length != data.length) {
			throw new IllegalArgumentException("Size of width array does not match table size");
		}
		this.columnWidths = widths;

		return this;
	}
	
	public Table setColumnFills(String[] fills) {
		if (fills.length != data.length) {
			throw new IllegalArgumentException("Size of fills array does not match table size");
		}
		this.cellFills = fills;

		return this;
	}
	
//	public void setColumnStyles(String[] styles) {
//		this.cellStyles = styles;
//	}
	
	/**
	 * Adds the table to the document and for convenience returns the document builder
	 */
	public DOCXBuilder addToDocument(MainDocumentPart main) {
		if (data == null) {
			throw new IllegalArgumentException("Must supply data before adding to document");
		}
		if (columnWidths == null) {
			throw new IllegalArgumentException("Must supply columns widths before adding to document");
		}
		
		TblPr tblPr = factory.factory.createTblPr();
		tblPr.setJc(factory.docx4jCreateJustification("left"));
		tblPr.setTblInd(factory.docx4jCreateTblIndent(indent));
		
		TblBorders borders = factory.factory.createTblBorders();
		borders.setTop(factory.docx4jCreateBorderDefinition(borderCSS));
		borders.setLeft(factory.docx4jCreateBorderDefinition(borderCSS));
		borders.setRight(factory.docx4jCreateBorderDefinition(borderCSS));
		borders.setBottom(factory.docx4jCreateBorderDefinition(borderCSS));
		borders.setInsideH(factory.docx4jCreateBorderDefinition(borderCSS));
		borders.setInsideV(factory.docx4jCreateBorderDefinition(borderCSS));
		tblPr.setTblBorders(borders);
		tblPr.setTblCellMar(factory.docx4jCreateCellMargin(cellMargin));
        
        Tbl table = factory.factory.createTbl();
        table.setTblPr(tblPr);
        
        TblGrid tblGrid = factory.factory.createTblGrid();
        table.setTblGrid(tblGrid);

        for (int i=0 ; i<cols; i++) {
        	TblGridCol gridCol = factory.factory.createTblGridCol();
        	gridCol.setW(BigInteger.valueOf(columnWidths[i]));
        	tblGrid.getGridCol().add(gridCol);
        }
        
        //Apply defaults to fills...
        String[] fill = new String[cols];
        Arrays.fill(fill, "auto");

        if (cellFills != null) {
        	for (int i=0; i<fill.length; i++) {
        		if (cellFills[i] != null) {
        			fill[i] = cellFills[i];
        		}
        	}
        }
        
        
        //Apply defaults to styles...
        String[] style = new String[cols];
        Arrays.fill(style, null);
        if (cellStyles != null) {
        	for (int i=0; i<cellStyles.length; i++) {
        		if (cellStyles[i] != null) {
        			style[i] = cellStyles[i];
        		}
        	}
	    }
        
        for (int row = 0; row<rows; row++) {
        	Tr tr = factory.factory.createTr();
        	TrPr trpr = factory.factory.createTrPr();
        	tr.setTrPr(trpr);
        	if (rowHeights != null && rowHeights[row] != null) {
        		CTHeight height = factory.factory.createCTHeight();
        		height.setHRule(STHeightRule.AT_LEAST);
        		height.setVal(new BigInteger(rowHeights[row].toString()));
        		trpr.getCnfStyleOrDivIdOrGridBefore().add(factory.factory.createCTTrPrBaseTrHeight(height));
        	}
        	table.getContent().add(tr);
        	
        	for (int col = 0; col < cols; col++) {
        		Tc tc = factory.factory.createTc();
        		tr.getContent().add(tc);
        		
        		TcPr tcpr = factory.factory.createTcPr();
        		tc.setTcPr(tcpr);

        		Integer span = hMerges.get(new Point(col, row));
        		int width = columnWidths[col];
        		if (span != null) {
        			for (int i=1; i<span; i++) {
        				width += columnWidths[col + i];
        			}
        			GridSpan gspan = factory.factory.createTcPrInnerGridSpan();
        			gspan.setVal(BigInteger.valueOf(span));
        			tcpr.setGridSpan(gspan);
        		}
        		TblWidth w = factory.factory.createTblWidth();
                w.setW(BigInteger.valueOf(width));
                tcpr.setTcW(w);
        		
        		CTShd shade = factory.factory.createCTShd();
        		shade.setFill(fill[col]);
        		tcpr.setShd(shade);
        		
        		if (this.borders != null && this.borders[col][row] != null) {
        			TcBorders tcborders = factory.factory.createTcPrInnerTcBorders();
        			
        			tcborders.setBottom(factory.docx4jCreateBorderDefinition(this.borders[col][row]));
        			tcborders.setRight(factory.docx4jCreateBorderDefinition(this.borders[col][row]));
        			tcborders.setLeft(factory.docx4jCreateBorderDefinition(this.borders[col][row]));
        			tcborders.setTop(factory.docx4jCreateBorderDefinition(this.borders[col][row]));

        			tcpr.setTcBorders(tcborders);
        		}
        		String content = data[col][row];
        		if (content == null)
        			content = "";
        		for (String para : content.split("[\\r\\n]+")) {
	        		factory.createParagraph(style[col]).addFromHTML(para).addTo(tc);
        		}
        		
        		if (span != null) {
        			col += span-1;
        		}
        	}	
        }
 
        main.addObject(table);
        
        return factory;
	}

	public Table addHorizontalMerge(int row, int startCol, int colSpan) {
		if (colSpan == 1)
			return this;
		//Check data
		for (int i = 1; i<colSpan; i++) {
			if (data[startCol + i][row] != null) {
				throw new IllegalArgumentException("Merge would overwrite data in row " + row + " column " + (startCol + i) + ".  Please ensure this cell is null");
			}
		}
		
		hMerges.put(new Point(startCol, row), colSpan);
		
		return this;
	}

	public Table setColumnStyles(String[] styles) {
		if (styles.length != data.length) {
			throw new IllegalArgumentException("Size of styles array does not match table size");
		}
		this.cellStyles = styles;
		
		return this;
	}

	public Table setRowHeights(Integer[] heights) {
		if (heights.length != data[0].length) {
			throw new IllegalArgumentException("Size of heights array does not match table size");
		}
		this.rowHeights = heights;
		
		return this;

	}

	public Table setTableBorders(String css) {
		borderCSS = css;
		
		return this;
	}

	public Table setCellBorders(int i, int j, String css) {
		if (borders == null) {
			borders = new String[data.length][data[0].length];
		}
		borders[i][j] = css;
		
		return this;
	}

}
