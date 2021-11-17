package com.jumbletree.docx5j.xlsx.trials;

import java.awt.Color;
import java.awt.Desktop;
import java.io.File;
import java.io.IOException;
import java.util.function.Consumer;

import javax.xml.bind.JAXBException;

import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.exceptions.InvalidFormatException;
import org.xlsx4j.sml.STBorderStyle;
import org.xlsx4j.sml.STHorizontalAlignment;
import org.xlsx4j.sml.STSheetViewType;
import org.xlsx4j.sml.STVerticalAlignment;

import com.jumbletree.docx5j.xlsx.XLSXFile;
import com.jumbletree.docx5j.xlsx.builders.CellBuilder;
import com.jumbletree.docx5j.xlsx.builders.RowBuilder;
import com.jumbletree.docx5j.xlsx.builders.SheetViewBuilder;
import com.jumbletree.docx5j.xlsx.builders.StyleBuilder.Edge;
import com.jumbletree.docx5j.xlsx.builders.TextModifierBuilder;
import com.jumbletree.docx5j.xlsx.builders.WorksheetBuilder;

public class Demo2 {

	public static void main(String[] args) throws Docx4JException, JAXBException, IOException {
		String org = "SkyNet Support Services";
		String date = "September 2021";
		
		XLSXFile file = createFile();
	
		WorksheetBuilder sheet = createSheet(file);
		
		addPageHeader(org, date, sheet);
		
		addHeaderRow(sheet, cell->{
			cell.multiStyleText()
				.add("Energy", TextModifierBuilder.newBuilder().withColor(Color.RED).build())
				.add(" Consumption (kWh)", TextModifierBuilder.newBuilder().withColor(Color.BLACK).build());
		}, true);
				
		//Header row - no data
		addTableHeaderLine(sheet, "Council Direcorate");
		addTableHeaderLine(sheet, "Electricity");
		
		addDataLine(sheet, "Infrastructure Services", 873676., 755723., 1801816, 1351344, 11028495, false, false);
		addDataLine(sheet, "City Services", 35358, 52912, 65622, 85439, 305053, false, false);
		addDataLine(sheet, "Total Electricity Consumption kWh", 1345932, 1194253, 2754577, 2248546, 15174621, true, false);
		
		addTableHeaderLine(sheet, "");

		addDataLine(sheet, "Total TOU Electricity Consumption kWh", 1166478, 1065927, 2467799, 1966203, 13626005, false, true);
		addDataLine(sheet, "Total Gas Consumption kWh", 253093, 313141, 543007, 613014, 2837082, false, true);

		addTableHeaderLine(sheet, "");
		addDataLine(sheet, "Total Energy Cost $", 267442, 287241, 577528, 577250, 2797194, true, false);
		addTableHeaderLine(sheet, "");
		addHeaderRow(sheet, "Energy Cost ($) - excl GST", false);
		addTableHeaderLine(sheet, "");

		addDataLine(sheet, "Average TOU c/kWh", 18.07, 24.38, 18.09, false, true, false);
		addDataLine(sheet, "Unaverage TOU c/kWh", 23.48, 18.47, 17.39, true, false, false);
		addDataLine(sheet, "Average c/kWh", 23.18, 18.37, 17.39, false, false, true);

		
		File out = new File("test1.xlsx");
		file.save(out);
		Desktop.getDesktop().open(out);

	}

	private static void addPageHeader(String org, String date, WorksheetBuilder sheet) throws Docx4JException {
		sheet.skipRow()
			.nextRow()
				.cells(9)
					.style("header-area")
					.row()
				.sheet()
			.nextRow()
				.height(29, true)
				.mergeCells(0, 5)
				.mergeCells(6, 8, 2)
				.nextCell()
					.style("header")
					.value(org + " Energy Dashboard")
					.row()
				//Skip cells that are under the merge
				.skipCells(5)
				.nextCell()
					.style("header")
					.value(date)
					.row()
				.sheet()
			.nextRow()
				.height(29, true)
				.mergeCells(0, 5)
				.nextCell()
					.style("header")
					.value("By Directorate and  Energy Type")
					.row()
				.sheet()
			.nextRow()
				.cells(9)
					.style("header-area")
					.row()
				.sheet()
			.skipRow();
	}

	private static WorksheetBuilder createSheet(XLSXFile file) throws Docx4JException {
		WorksheetBuilder sheet = file.getWorkbookBuilder()
			.getSheet(0)
			.addColumnDefinition(41.77, false)
			.addColumnDefinition(12.62, false, 2,9);
		
		SheetViewBuilder.newBuilder(STSheetViewType.NORMAL)
				 .withProperty(SheetViewBuilder.SHOW_GRID_LINES, false)
				 .addTo(sheet);
		return sheet;
	}

	private static XLSXFile createFile() throws InvalidFormatException, JAXBException {
		XLSXFile file = new XLSXFile();

		//"Normal" styles
		file.createStyle()
			//"Normal"
			.withFont("Calibri", 11, Color.black, false, false)
			.withFormat("#,##0")
			.installAs("norm")
			.copy()
			//"Italic"
			.withFont("Calibri", 11, Color.black, false, true)
			.installAs("norm-italic")
			.copy()
			//"Bold"
			.withFont("Calibri", 11, Color.black, true, false)
			.installAs("norm-bold")
			.copy()
			//Bottom
			.withBorder(STBorderStyle.MEDIUM, Edge.BOTTOM)
			.installAs("norm-bottom")
			.copy()
			.withFormat("0.00")
			.installAs("norm-twodp-bottom")
			.copy()
			//2dp
			.clearBorders()
			.withFont("Calibri", 11, Color.black, false, false)
			.installAs("norm-twodp")
			.copy()
			//"Italic"
			.withFont("Calibri", 11, Color.black, false, true)
			.installAs("norm-twodp-italic")
			.copy()
			//"Bold"
			.withFont("Calibri", 11, Color.black, true, false)
			.installAs("norm-twodp-bold");
			
		//Table header styles (static)
		file.createStyle()
			//Table header, left
			.withFill(Color.decode("0xF2F2F2"))
			.withBorder(STBorderStyle.MEDIUM, Edge.LEFT)
			.installAs("table-header-left")
			.copy()
			//Table header, top
			.withBorder(STBorderStyle.MEDIUM, Edge.TOP)
			.installAs("table-header-top-left")
			.copy()
			//Table header, top
			.withBorder(STBorderStyle.MEDIUM, Edge.TOP)
			.withAlignment(STHorizontalAlignment.CENTER, STVerticalAlignment.BOTTOM, true)
			.installAs("table-header-top")
			.copy()
			//Table header, now border
			.clearBorders()
			.installAs("table-header")
			.copy()
			//Table header, top right
			.withBorder(STBorderStyle.MEDIUM, Edge.TOP_RIGHT)
			.installAs("table-header-top-right")
			.copy()
			//Table header, right
			.withBorder(STBorderStyle.MEDIUM, Edge.RIGHT)
			.withAlignment(STHorizontalAlignment.CENTER, STVerticalAlignment.BOTTOM, true)
			.installAs("table-header-right");

		//Table header styles, dynamic
		file.createStyle()
			//Table header, left, bold
			.withFill(Color.decode("0xF2F2F2"))
			.withAlignment(STHorizontalAlignment.LEFT, STVerticalAlignment.BOTTOM)
			.withBorder(STBorderStyle.MEDIUM, Edge.LEFT)
			.withFont("Calibri", 11, Color.black, true, false)
			.installAs("table-left-bold")
			.copy()
			//Table header, left, italic
			.withFont("Calibri", 11, Color.black, false, true)
			.installAs("table-left-italic")
			.copy()
			//Table header, left, no bold
			.withFont("Calibri", 11, Color.black, false, false)
			.installAs("table-left")
			.copy()
			//Table header, bottom left
			.withBorder(STBorderStyle.MEDIUM, Edge.BOTTOM_LEFT)
			.withFont("Calibri", 11, Color.black, true, false)
			.withAlignment(STHorizontalAlignment.LEFT, STVerticalAlignment.BOTTOM)
			.installAs("table-left-bottom")
			.copy()
			//Table bottom
			.clearFill()
			.withBorder(STBorderStyle.MEDIUM, Edge.BOTTOM)
			.withFont("Calibri", 11, Color.black, true, false)
			.installAs("norm-bottom");

		//Right margin styles
		file.createStyle()
			//Table right
			.withBorder(STBorderStyle.MEDIUM, Edge.RIGHT)
			.withAlignment(STHorizontalAlignment.RIGHT, STVerticalAlignment.BOTTOM)
			.withFont("Calibri", 11, Color.black, false, false)
			.withFormat("#,##0")
			.installAs("table-right")
			.copy()
			//Table right - bold
			.withFont("Calibri", 11, Color.black, true, false)
			.installAs("table-right-bold")
			.copy()
			//Table right - italic
			.withFont("Calibri", 11, Color.black, false, true)
			.installAs("table-right-italic")
			.copy()
			//Table right - 2dp italic
			.withFormat("0.00")
			.installAs("table-right-twodp-italic")
			.copy()
			//Table right - 2dp normal
			.withFont("Calibri", 11, Color.black, false, false)
			.installAs("table-right-twodp")
			.copy()
			//Table right - 2dp bold
			.withFont("Calibri", 11, Color.black, true, false)
			.installAs("table-right-twodp-bold")
			.copy()
			//Table right - 2dp bold
			.withBorder(STBorderStyle.MEDIUM, Edge.BOTTOM_RIGHT)
			.installAs("table-right-twodp-bottom");
		
		//Page header styles
		file.createStyle()
			//Blue header, no text
			.clearBorders()
			.withFill(Color.decode("0x1065b7"))
			.installAs("header-area")
			.copy()
			//Blue header text
			.withFont("Calibri", 22, Color.white, true, false)
			.withAlignment(STHorizontalAlignment.CENTER, STVerticalAlignment.CENTER)
			.installAs("header");
		
		//Negative variance styles
		file.createStyle()
			//Good, bold
			.withFill(Color.decode("0xC6EFCE"))
			.withAlignment(STHorizontalAlignment.CENTER, STVerticalAlignment.CENTER)
			.withFormat("0%")
			.withFont("Calibri", 11, Color.decode("0x006100"), true, false)
			.installAs("negative-variance-bold")
			.copy()
			.withBorder(STBorderStyle.MEDIUM, Edge.BOTTOM)
			.installAs("negative-variance-bottom")
			.copy()
			//Good, italic
			.clearBorders()
			.withFont("Calibri", 11, Color.decode("0x006100"), false, true)
			.installAs("negative-variance-italic")
			.copy()
			//Good, no bold
			.withFont("Calibri", 11, Color.decode("0x006100"), false, false)
			.installAs("negative-variance");
	
		//Positive variance styles
		file.createStyle()
			//Good, bold
			.withFill(Color.decode("0xFFC7CE"))
			.withAlignment(STHorizontalAlignment.CENTER, STVerticalAlignment.CENTER)
			.withFormat("0%")
			.withFont("Calibri", 11, Color.decode("0x9C0006"), true, false)
			.installAs("positive-variance-bold")
			.copy()
			.withBorder(STBorderStyle.MEDIUM, Edge.BOTTOM)
			.installAs("positive-variance-bottom")
			.copy()
			//Good, italic
			.clearBorders()
			.withFont("Calibri", 11, Color.decode("0x9C0006"), false, true)
			.installAs("positive-variance-italic")
			.copy()
			//Good, no bold
			.withFont("Calibri", 11, Color.decode("0x9C0006"), false, false)
			.installAs("positive-variance");
		return file;
	}

	private static void addHeaderRow(WorksheetBuilder sheet, String header, boolean first) throws Docx4JException {
		addHeaderRow(sheet, cell->cell.value(header), first);
	}
	
	private static void addHeaderRow(WorksheetBuilder sheet, Consumer<CellBuilder> header, boolean first) throws Docx4JException {
		CellBuilder cell = sheet
			.nextRow()
				//.height(32, true)
				.nextCell()
					.style(first ? "table-header-top-left" : "table-header-left");
					header.accept(cell);
					cell.row()
				.cells(7)
					.style(first ? "table-header-top" : "table-header")
					.values("This Month Last Year", "This Month Actual", "% Var", "YTD Last Year", "YTD Actual", "% Var", "Forecast 2021/2022")
					.row()
				.nextCell()
					.style(first ? "table-header-top-right" : "table-header-right")
					.value("Last Year 2020/2021");
	}

	private static void addDataLine(WorksheetBuilder sheet, 
			String header,
			double lastMonth, double thisMonth, double lastYTD, double thisYTD, double lastFY, 
			boolean bold, boolean italic) throws Docx4JException {
		addDataLine(sheet, header, lastMonth, thisMonth, lastYTD, thisYTD, lastFY, bold, italic, false, "");
	}
		
	private static void addDataLine(WorksheetBuilder sheet, 
			String header,
			double lastYTD, double thisYTD, double lastFY, 
			boolean bold, boolean italic, boolean last) throws Docx4JException {
		addDataLine(sheet, header, null, null, lastYTD, thisYTD, lastFY, bold, italic, last, "-twodp");
	}
		
	private static void addDataLine(WorksheetBuilder sheet, 
			String header,
			Double lastMonth, Double thisMonth, double lastYTD, double thisYTD, double lastFY, 
			boolean bold, boolean italic, boolean last, String formatSuffix) throws Docx4JException {

		double ytdVar = (thisYTD - lastYTD) / lastYTD;
		
		double forecast = lastFY * (1+ytdVar);
		
		String styleSuffix  = "";
		if (last) {
			styleSuffix = "-bottom";
		} else if (bold) {
			styleSuffix = "-bold";
		} else if (italic) {
			styleSuffix = "-italic";
		}

		String norm = "norm" + formatSuffix + styleSuffix;
		
		RowBuilder row = sheet.nextRow();
		
		row
			.nextCell()
				.style("table-left" + styleSuffix)
				.value(header);

		if (lastMonth == null || thisMonth == null) {
			row
				.cells(3)
					.style(norm)
					.row();
		} else {
			double monthVar = (thisMonth - lastMonth) / lastMonth;
			row
				.nextCell()
					.style(norm)
					.value(lastMonth)
					.row()
				.nextCell()
					.style(norm)
					.value(thisMonth)
					.row()
				.nextCell()
					.style((monthVar <= 0 ? "negative-variance" : "positive-variance") + styleSuffix)
					.value(monthVar);
		}
		row
			.nextCell()
				.style(norm)
				.value(lastYTD)
				.row()
			.nextCell()
				.style(norm)
				.value(thisYTD)
				.row()
			.nextCell()
				.style((ytdVar <= 0 ? "negative-variance" : "positive-variance") + styleSuffix)
				.value(ytdVar)
				.row()
			.nextCell()
				.style(norm)
				.value(forecast)
				.row()
			.nextCell()
				.style("table-right" + formatSuffix + styleSuffix)
				.value(lastFY);
					
				
	}

	private static void addTableHeaderLine(WorksheetBuilder sheet, String header) throws Docx4JException {
		sheet
			.nextRow()
				.nextCell()
					.style("table-header-left")
					.value(header)
					.row()
				.skipCells(7)
				.nextCell()
					.style("table-right");

	}

}
