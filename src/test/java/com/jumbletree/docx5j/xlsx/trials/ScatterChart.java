package com.jumbletree.docx5j.xlsx.trials;

import java.awt.Color;
import java.io.File;
import java.io.IOException;

import javax.xml.bind.JAXBException;

import org.docx4j.openpackaging.exceptions.Docx4JException;

import com.jumbletree.docx5j.xlsx.LineChart;
import com.jumbletree.docx5j.xlsx.LineProperties;
import com.jumbletree.docx5j.xlsx.Location;
import com.jumbletree.docx5j.xlsx.XLSXFile;
import com.jumbletree.docx5j.xlsx.XLSXRange;
import com.jumbletree.docx5j.xlsx.builders.WorkbookBuilder;

public class ScatterChart {

	public static void main(String[] args) throws Docx4JException, JAXBException, IOException {
		XLSXFile file = new XLSXFile();
		WorkbookBuilder factory = file.getWorkbookBuilder();

		file.createStyle().withFont("Calibri", 11, Color.black, true, false).installAs("bold");
		
		factory.setValue(0, 1, 0, "A Heading", "bold");

		for (int i=1; i<=3; i++) {
			factory.setValue(0, i, 3, "Column " + i);
		}

		for (int i=1; i<=3; i++) {
			factory.setValue(0, 0, 3+i, "Row " + i);
		}

		for (int i=1; i<=3; i++) {
			for (int j=1; j<=3; j++) {
				factory.setValue(0, j, 3+i, new Integer(i + "" + j).doubleValue());
			}
		}

		LineChart chart = factory.createLineChart();
		chart.setLegendPosition(Location.BOTTOM);
		chart.setCatRange(new XLSXRange("Sheet1", "A5", "A7"), false);
		chart.addSeries(new XLSXRange("Sheet1", "B4", "B4"), new XLSXRange("Sheet1", "B5", "B7"), new LineProperties(new Color(0xff, 0x42, 0x0e)), null);
		chart.addSeries(new XLSXRange("Sheet1", "C4", "C4"), new XLSXRange("Sheet1", "C5", "C7"), new LineProperties(new Color(0xff, 0x42, 0x0e)), null);
		chart.addSeries(new XLSXRange("Sheet1", "D4", "D4"), new XLSXRange("Sheet1", "D5", "D7"), new LineProperties(new Color(0xff, 0xd3, 0x20)), null);

		chart.setTitle("Chart Twitle");
		chart.setXAxisLabel("Month");
		chart.setYAxisLabel("Cost ($ excl GST)");
		chart.create(new XLSXRange("Sheet1", "E5", "L30"));

		//		ScatterChart chart2 = factory.createScatterChart();
		//		chart2.addSeries("snob", new XLSXRange("Sheet1", "B5", "B7"), new XLSXRange("Sheet1", "C5", "C7"), new LineProperties(new Color(0xff, 0x42, 0x0e)), null);
		//		chart2.addSeries(new XLSXRange("Sheet1", "B1", "B1"), new XLSXRange("Sheet1", "B5", "B7"), new XLSXRange("Sheet1", "D5", "D7"), new LineProperties(new Color(0xff, 0xd3, 0x20)), null);
		//		
		//		chart2.create(new XLSXRange("Sheet1", "J5", "N10"));

		File out = new File("C:/Users/matchami/Desktop/trial1.xlsx");
		file.save(out);
//		Desktop.getDesktop().open(file);
	}
}
