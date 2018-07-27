package com.jumbletree.docx5j.xlsx;

import java.util.ArrayList;

import org.docx4j.dml.chart.CTAxDataSource;
import org.docx4j.dml.chart.CTChart;
import org.docx4j.dml.chart.CTDLbls;
import org.docx4j.dml.chart.CTPlotArea;
import org.docx4j.dml.chart.CTScatterChart;
import org.docx4j.dml.chart.CTScatterSer;
import org.docx4j.dml.chart.CTScatterStyle;
import org.docx4j.dml.chart.STAxPos;
import org.docx4j.dml.chart.STScatterStyle;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.exceptions.InvalidFormatException;
import org.docx4j.openpackaging.packages.SpreadsheetMLPackage;
import org.docx4j.openpackaging.parts.SpreadsheetML.WorksheetPart;

import com.jumbletree.docx5j.xlsx.builders.WorkbookBuilder;

public class ScatterChart extends XLSXChart {
	private ArrayList<XLSXRange> xseries;
	private ArrayList<XLSXRange> yseries;
	private ArrayList<SeriesTitle> seriesTitles;
	private ArrayList<LineProperties> seriesLines;
	private ArrayList<MarkerProperties> seriesMarkers;
	
	public ScatterChart(WorksheetPart sheet, SpreadsheetMLPackage pkg, WorkbookBuilder factory) throws InvalidFormatException {
		super(sheet, pkg, factory);
		this.xseries = new ArrayList<XLSXRange>();
		this.yseries = new ArrayList<XLSXRange>();
		this.seriesTitles = new ArrayList<SeriesTitle>();
		this.seriesLines = new ArrayList<>();
		this.seriesMarkers = new ArrayList<>();
	}
		 
	public void addSeries(String title, XLSXRange xdata, XLSXRange ydata, LineProperties line, MarkerProperties marker) {
		this.xseries.add(xdata);
		this.yseries.add(ydata);
		this.seriesTitles.add(new SeriesTitle(false, title));
		this.seriesLines.add(line);
		this.seriesMarkers.add(marker);
	}
	
	public void addSeries(XLSXRange title, XLSXRange xdata, XLSXRange ydata, LineProperties line, MarkerProperties marker) {
		this.xseries.add(xdata);
		this.yseries.add(ydata);
		this.seriesTitles.add(new SeriesTitle(true, title.singleCellAbsoluteReference()));
		this.seriesLines.add(line);
		this.seriesMarkers.add(marker);
	}
	
	protected void createChart(CTChart chart, CTPlotArea ctp) throws Docx4JException {
		//And now the line chart.... 
		CTScatterChart scatterChart = new CTScatterChart();
		
		//TODO create other marker types
		CTScatterStyle style = new CTScatterStyle();
		//LINE_MARKER means line only
		style.setVal(STScatterStyle.LINE_MARKER);
		scatterChart.setScatterStyle(style);
		scatterChart.setVaryColors(createBoolean(false));
	      
	      for (int seriesIndex=0; seriesIndex<xseries.size(); seriesIndex++) {
	    	  CTScatterSer serie = new CTScatterSer();
	    	  scatterChart.getSer().add(serie);
	    	  
	    	  //Metadata
	    	  serie.setIdx(createUnsignedInt(seriesIndex));
	    	  serie.setOrder(createUnsignedInt(seriesIndex));
	    	  
	    	  //Series Title
	    	  serie.setTx(createSeriesTitle(seriesTitles.get(seriesIndex)));
	    	  
	    	  //Line and marker
	    	  serie.setSpPr(createLineProperties(seriesLines.get(seriesIndex)));
	    	  //no markers for Line, but we'll need to add them in when the others are included
	    	  //serie.setMarker(createMarkerProperties(seriesMarkers.get(seriesIndex)));

	    	  CTAxDataSource data = new CTAxDataSource();
	    	  data.setNumRef(createNumRef(xseries.get(seriesIndex)));
	    	  serie.setXVal(data);
	    	  serie.setYVal(createNumDataSource(yseries.get(seriesIndex)));
	    	  serie.setSmooth(createBoolean(false));
	    	  //TODO series labels
	      }
	      
	      CTDLbls lbls = new CTDLbls();
	      lbls.setShowLegendKey(createBoolean(false));
	      lbls.setShowVal(createBoolean(false));
	      lbls.setShowCatName(createBoolean(false));
	      lbls.setShowSerName(createBoolean(false));
	      lbls.setShowPercent(createBoolean(false));
	      lbls.setShowBubbleSize(createBoolean(false));
	      scatterChart.setDLbls(lbls);
	      
	      int cataxId = (int)Math.round(Math.random() * 10000000);
	      int valaxId = (int)Math.round(Math.random() * 10000000);
	      scatterChart.getAxId().add(createUnsignedInt(cataxId));
	      scatterChart.getAxId().add(createUnsignedInt(valaxId));
	      //objs.add(objBar);
	      ctp.getAreaChartOrArea3DChartOrLineChart().add(scatterChart);

	      setupAxes(ctp, cataxId, valaxId, true, 0);
	}

}
