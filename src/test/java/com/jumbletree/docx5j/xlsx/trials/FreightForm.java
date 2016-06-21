package com.jumbletree.docx5j.xlsx.trials;

import java.awt.Color;
import java.awt.Desktop;
import java.io.File;
import java.io.IOException;

import javax.xml.bind.JAXBException;

import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.xlsx4j.sml.STSheetViewType;

import com.jumbletree.docx5j.xlsx.SheetFormat;
import com.jumbletree.docx5j.xlsx.View;
import com.jumbletree.docx5j.xlsx.XLSXFile;
import com.jumbletree.docx5j.xlsx.XLSXRange;

public class FreightForm {

	public static void main(String[] args) throws JAXBException, Docx4JException, IOException {
		XLSXFile file = new XLSXFile();
		
		file.createStyle().withFont("Calibri", 14, Color.black, true, false).installAs("title");
		//First sheet is created automatically
		file.getWorkbookBuilder()
			.getSheet(0)
			.setTabColor(new Color(0xFF, 0xC0, 0x00))
			.setName("Freight")
			.setDimension(new XLSXRange("A1", "G37"))
			.addView(new View(0, 80, 100, 80, STSheetViewType.PAGE_BREAK_PREVIEW, true))
			.setFormat(new SheetFormat(15))
			.addColumnDefinition(21.28515625, false)
			.addColumnDefinition(54.85546875, true)
			.addColumnDefinition(19, false)
			.addColumnDefinition(19.85546875, true)
			.addColumnDefinition(15.5703125, true)
			.addColumnDefinition(14.85546875, true)
			.nextRow()
				.setHeight(18.75)
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
			;
				

		
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
		
		File out = new File("C:\\Users\\matchami\\Desktop\\test.xlsx");
		file.save(out);
		Desktop.getDesktop().open(out);
	}

	private static String getDateRange() {
		return "APRIL - JUNE 2016";
	}

}
