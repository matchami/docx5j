package com.jumbletree.docx5j.xlsx.trials;

import java.awt.Color;
import java.awt.Desktop;
import java.io.File;
import java.io.IOException;
import java.util.Arrays;

import javax.xml.bind.JAXBException;

import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.xlsx4j.sml.STBorderStyle;
import org.xlsx4j.sml.STHorizontalAlignment;
import org.xlsx4j.sml.STOrientation;
import org.xlsx4j.sml.STPatternType;
import org.xlsx4j.sml.STSheetViewType;
import org.xlsx4j.sml.STVerticalAlignment;

import com.jumbletree.docx5j.xlsx.CommentPosition;
import com.jumbletree.docx5j.xlsx.SheetFormat;
import com.jumbletree.docx5j.xlsx.View;
import com.jumbletree.docx5j.xlsx.XLSXFile;
import com.jumbletree.docx5j.xlsx.XLSXRange;
import com.jumbletree.docx5j.xlsx.builders.WorksheetBuilder;

public class FreightForm {

	public static void main(String[] args) throws JAXBException, Docx4JException, IOException {
		XLSXFile file = new XLSXFile();
		
		file.createStyle().withFont("Calibri", 14, Color.black, true, false).installAs("title");
		file.createStyle()
			.withFont("Calibri", 12, Color.black, true, false)
			.withBorder(STBorderStyle.THIN, Color.black)
			.withAlignment(STHorizontalAlignment.CENTER, STVerticalAlignment.CENTER)
			.installAs("th");
		file.createStyle()
			.withFont("Calibri", 11, new Color(0, 176, 240), false, true)
			.withBorder(STBorderStyle.THIN, Color.black)
			.withAlignment(STHorizontalAlignment.CENTER, null)
			.withFormat(3)
			.installAs("eg")
			.copy()
			.withFormat(0)
			.installAs("egDecimal")
			.copy()
			.withFont("Calibri", 11, Color.black, false, false)
			.installAs("table");
		
		file.createStyle()
			.withFont("Calibri", 11, Color.black, true, false)
			.withBorder(STBorderStyle.MEDIUM, Color.black, STBorderStyle.THIN, Color.black, STBorderStyle.MEDIUM, Color.black, STBorderStyle.THIN, Color.black)
			.withFill(new Color(230, 184, 183), new Color(230, 184, 183), STPatternType.SOLID)
			.installAs("month");
		
		//First sheet is created automatically
		WorksheetBuilder sheet = file.getWorkbookBuilder()
			.getSheet(0)
			.setTabColor(new Color(0xFF, 0xC0, 0x00))
			.setName("Freight")
			.setDimension(new XLSXRange("A1", "G37"))
			.addView(new View(0, 100, 100, 100, STSheetViewType.PAGE_BREAK_PREVIEW, true))
			.setFormat(new SheetFormat(15))
			.addColumnDefinition(21.28515625, false)
			.addColumnDefinition(54.85546875, true)
			.addColumnDefinition(19, false)
			.addColumnDefinition(19.85546875, true)
			.addColumnDefinition(15.5703125, true)
			.addColumnDefinition(14.85546875, true)
			.nextRow()
				.height(18.75)
				.nextCell()
					.style("title")
					.value("FREIGHT DATA:")
					.row()
				.nextCell()
					.style("title")
					.value(getDateRange())
					.row()
				.addExplicitSpan(1,6)
				.sheet()
			.nextRow()
				.height(32.25, true)
				.cells(6)
					.style("th")
					.values("DATE", "FROM", "TO", "WEIGHT (kgs)", "DISTANCE (kms)", "CLUB")
					.row()
				.addExplicitSpan(1,6)
				.sheet()
			.nextRow()
				.height(15.75, true)
				.cells(4)
					.style("eg")
					.values("E.g. 10/02/2015", "LMNZ", "LM CHC")
					.style("egDecimal")
					.values(2.5)
					.row()
				.nextCell()
					.style("eg")
					.comment("Kelly Reuben", "See next tab for distances.", new CommentPosition(3, 101, 2, 8, 4, 94, 5, 0))
					.value(1069)
					.row()
				.nextCell()
					.style("eg")
					.value("LMNZ")
					.row()
				.addExplicitSpan(1,6)
				.sheet();
		
		for (String month : Arrays.asList("April 2015", "May 2015", "June 2015")) {
			sheet.nextRow()
				.height(15.75)
				.cells(6)
					.style("month")
					.values(month)
					.row()
				.sheet();
		
			for (int i=0; i<5; i++) {
				sheet.nextRow()
					.height(15.75)
					.cells(6)
						.style("table")
						.row();
			}
		}

		sheet			
			.setPageMargns(.3, .3, .75, .7, .75, .7)
			.setPageSetup(FreightForm.class.getResourceAsStream("printerSettings1.bin"), STOrientation.LANDSCAPE, 78, 9)
			.addRowBreak(20);

		
		file.getWorkbookBuilder()
			.appendSheet()
			.setTabColor(new Color(0x00, 0x70, 0xC0))
			.setName("Distances")
			.setDimension(new XLSXRange("B1", "D15"))
			.addView(new View(0))
			.setFormat(new SheetFormat(15))
			.addColumnDefinition(10.85546875, false, 2)
			.addColumnDefinition(17.7109375, true)
			.addColumnDefinition(11, false)
			;
		
		File out = new File("test1.xlsx");
		file.save(out);
		Desktop.getDesktop().open(out);
	}

	private static String getDateRange() {
		return "APRIL - JUNE 2016";
	}

}
