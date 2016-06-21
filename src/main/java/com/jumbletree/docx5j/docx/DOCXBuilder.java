package com.jumbletree.docx5j.docx;

import java.awt.Insets;
import java.math.BigInteger;

import org.docx4j.wml.CTBorder;
import org.docx4j.wml.CTTblCellMar;
import org.docx4j.wml.Jc;
import org.docx4j.wml.JcEnumeration;
import org.docx4j.wml.ObjectFactory;
import org.docx4j.wml.STBorder;
import org.docx4j.wml.TblWidth;

public class DOCXBuilder {

	ObjectFactory factory;
	private boolean cuedPageBreak = false;;

	public DOCXBuilder() {
		factory = new ObjectFactory();
	}
	
	public Paragraph createParagraph() {
		return createParagraph(null);
	}
	
	public Paragraph createParagraph(String style) {
		Paragraph p = new Paragraph(this, style, cuedPageBreak);
		cuedPageBreak = false;
		return p;
	}
	
	public Table createTable(String[][] data) {
		if (data.length == 0) {
			throw new IllegalArgumentException("Table would have zero size");
		}
		return new Table(this, data);
	}
	
	public Style createStyle(String css) {
		return new Style(css, factory);
	}

	public DOCXBuilder pageBreak() {
		this.cuedPageBreak  = true;
		return this;
	}

	public Jc docx4jCreateJustification(String justify) {
		Jc jc = factory.createJc();
		jc.setVal(JcEnumeration.valueOf(justify.toUpperCase()));
		return jc;
	}

	public TblWidth docx4jCreateTblIndent(int indent) {
		TblWidth width = factory.createTblWidth();
		width.setType("dxa");
		width.setW(BigInteger.valueOf(indent));
		return width;
	}

	public CTBorder docx4jCreateBorderDefinition(String borderCSS) {
		CTBorder border = factory.createCTBorder();
		if (borderCSS == null || borderCSS.length() == 0 || borderCSS.equalsIgnoreCase("none")) {
			border.setVal(STBorder.NONE);
		} else {
			String[] bits = borderCSS.split("\\s");
			
			//Expect "style color size"
			if (bits[0].equals("solid")) {
				border.setVal(STBorder.SINGLE);
			} else {
				border.setVal(STBorder.valueOf(bits[0].toUpperCase()));
			}
			
			if (!bits[1].startsWith("#")) {
				throw new IllegalArgumentException("Please specify border color (" + bits[1] + ") as #hex value");
			}
			border.setColor(bits[1].substring(1));
			
			if (bits[2].endsWith("px")) {
				bits[2] = bits[2].substring(0, bits[2].length()-2);
			}
			border.setSz(BigInteger.valueOf(Integer.parseInt(bits[2]) * 2));
			border.setSpace(BigInteger.ZERO);
		}
		return border;
	}

	public CTTblCellMar docx4jCreateCellMargin(Insets cellMargin) {
		CTTblCellMar margin = factory.createCTTblCellMar();
		margin.setTop(docx4jCreateTblIndent(cellMargin.top));
		margin.setLeft(docx4jCreateTblIndent(cellMargin.left));
		margin.setRight(docx4jCreateTblIndent(cellMargin.right));
		margin.setBottom(docx4jCreateTblIndent(cellMargin.bottom));
		return margin;
	}
}
