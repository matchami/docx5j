package com.jumbletree.docx5j.xlsx;

import java.util.ArrayList;

import org.docx4j.dml.CTEffectList;
import org.docx4j.dml.CTLineProperties;
import org.docx4j.dml.CTNoFillProperties;
import org.docx4j.dml.CTShapeProperties;
import org.docx4j.dml.chart.CTChart;
import org.docx4j.dml.chart.CTChartLines;
import org.docx4j.dml.chart.CTDLblPos;
import org.docx4j.dml.chart.CTDLbls;
import org.docx4j.dml.chart.CTGapAmount;
import org.docx4j.dml.chart.CTGrouping;
import org.docx4j.dml.chart.CTLineChart;
import org.docx4j.dml.chart.CTLineSer;
import org.docx4j.dml.chart.CTPlotArea;
import org.docx4j.dml.chart.CTUpDownBar;
import org.docx4j.dml.chart.CTUpDownBars;
import org.docx4j.dml.chart.STDLblPos;
import org.docx4j.dml.chart.STGrouping;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.exceptions.InvalidFormatException;
import org.docx4j.openpackaging.packages.SpreadsheetMLPackage;
import org.docx4j.openpackaging.parts.SpreadsheetML.WorksheetPart;

public class LineChart extends XLSXChart {

	private ArrayList<XLSXRange> series;
	private ArrayList<SeriesTitle> seriesTitles;
	private ArrayList<LineProperties> seriesLines;
	private ArrayList<MarkerProperties> seriesMarkers;
	private XLSXRange cat;
	private boolean catIsNum;
	
	public LineChart(WorksheetPart sheet, SpreadsheetMLPackage pkg, XLSXFactory factory) throws InvalidFormatException {
		super(sheet, pkg, factory);
		this.series = new ArrayList<XLSXRange>();
		this.seriesTitles = new ArrayList<SeriesTitle>();
		this.seriesLines = new ArrayList<>();
		this.seriesMarkers = new ArrayList<>();
	}
	
	public void setCatRange(XLSXRange range, boolean isNum) {
		this.cat = range;
		this.catIsNum = isNum;
	}
		 
	public void addSeries(String title, XLSXRange data, LineProperties line, MarkerProperties marker) {
		this.series.add(data);
		this.seriesTitles.add(new SeriesTitle(false, title));
		this.seriesLines.add(line);
		this.seriesMarkers.add(marker);
	}
	
	public void addSeries(XLSXRange title, XLSXRange data, LineProperties line, MarkerProperties marker) {
		this.series.add(data);
		this.seriesTitles.add(new SeriesTitle(true, title.singleCellAbsoluteReference()));
		this.seriesLines.add(line);
		this.seriesMarkers.add(marker);
	}
	
	protected void createChart(CTChart chart, CTPlotArea ctp) throws Docx4JException {
		//And now the line chart.... 
	      CTLineChart lineChart = new CTLineChart();
	      CTGrouping grouping = new CTGrouping();
	      grouping.setVal(STGrouping.STANDARD);
	      lineChart.setGrouping(grouping);
	      lineChart.setVaryColors(createBoolean(true));
	      
	      for (int seriesIndex=0; seriesIndex<series.size(); seriesIndex++) {
	    	  CTLineSer serie = new CTLineSer();
	    	  lineChart.getSer().add(serie);
	    	  
	    	  //Metadata
	    	  serie.setIdx(createUnsignedInt(seriesIndex));
	    	  serie.setOrder(createUnsignedInt(seriesIndex));
	    	  
	    	  //Series Title
	    	  serie.setTx(createSeriesTitle(seriesTitles.get(seriesIndex)));
	    	  
	    	  //Line and marker
	    	  serie.setSpPr(createLineProperties(seriesLines.get(seriesIndex)));
	    	  
			  serie.setDLbls(createEmptyDataLabels());
	    	  
	    	  serie.setMarker(createMarkerProperties(seriesMarkers.get(seriesIndex)));
			  serie.setSmooth(createBoolean(false));

	    	  serie.setVal(createNumDataSource(series.get(seriesIndex)));
	    	  
	    	  //Categoryes
	    	  if (cat != null) {
	    		  serie.setCat(createAxDataSource(cat, catIsNum));
	    	  }
	      }
	      
	      lineChart.setDLbls(createEmptyDataLabels());
	      
	      CTChartLines lines = new CTChartLines();
	      CTShapeProperties sppr = new CTShapeProperties();
	      CTLineProperties ln = new CTLineProperties();
	      ln.setNoFill(new CTNoFillProperties());
	      sppr.setLn(ln);
	      lines.setSpPr(sppr);
//	      lineChart.setHiLowLines(lines);
	      
	      CTUpDownBars bars = new CTUpDownBars();
	      CTGapAmount width = new CTGapAmount();
	      width.setVal(150);
	      bars.setGapWidth(width);
	      bars.setUpBars(new CTUpDownBar());
	      bars.setDownBars(new CTUpDownBar());
//	      lineChart.setUpDownBars(bars);
//	      lineChart.setMarker(createBoolean(false));
	      lineChart.setSmooth(createBoolean(false));
	      
	      int cataxId = (int)Math.floor(Math.random() * 100000000);
	      int valaxId = (int)Math.floor(Math.random() * 100000000);
	      lineChart.getAxId().add(createUnsignedInt(cataxId));
	      lineChart.getAxId().add(createUnsignedInt(valaxId));
	      //objs.add(objBar);
	      ctp.getAreaChartOrArea3DChartOrLineChart().add(lineChart);

	      setupAxes(ctp, cataxId, valaxId);
	}

	protected CTDLbls createEmptyDataLabels() {
		CTDLbls dlbls = new CTDLbls();
		  
		  CTShapeProperties props = new CTShapeProperties();
		  props.setNoFill(new CTNoFillProperties());
		  CTLineProperties line = new CTLineProperties();
		  line.setNoFill(new CTNoFillProperties());
		  props.setLn(line);
		  props.setEffectLst(new CTEffectList());
		  dlbls.setSpPr(props);
		  CTDLblPos pos = new CTDLblPos();
		  pos.setVal(STDLblPos.R);
		  dlbls.setDLblPos(pos);
		  dlbls.setShowLegendKey(createBoolean(false));
		  dlbls.setShowVal(createBoolean(false));
		  dlbls.setShowCatName(createBoolean(false));
		  dlbls.setShowSerName(createBoolean(false));
		  dlbls.setShowPercent(createBoolean(false));
		  return dlbls;
	}
}
