package com.jumbletree.docx5j.xlsx;

import java.awt.Color;

import javax.xml.bind.JAXBElement;
import javax.xml.namespace.QName;

import org.docx4j.dml.CTEffectList;
import org.docx4j.dml.CTLineJoinRound;
import org.docx4j.dml.CTLineProperties;
import org.docx4j.dml.CTNoFillProperties;
import org.docx4j.dml.CTNonVisualDrawingProps;
import org.docx4j.dml.CTNonVisualGraphicFrameProperties;
import org.docx4j.dml.CTPoint2D;
import org.docx4j.dml.CTPositiveSize2D;
import org.docx4j.dml.CTRegularTextRun;
import org.docx4j.dml.CTSRgbColor;
import org.docx4j.dml.CTShapeProperties;
import org.docx4j.dml.CTSolidColorFillProperties;
import org.docx4j.dml.CTTextBody;
import org.docx4j.dml.CTTextBodyProperties;
import org.docx4j.dml.CTTextCharacterProperties;
import org.docx4j.dml.CTTextListStyle;
import org.docx4j.dml.CTTextParagraph;
import org.docx4j.dml.CTTextParagraphProperties;
import org.docx4j.dml.CTTransform2D;
import org.docx4j.dml.Graphic;
import org.docx4j.dml.GraphicData;
import org.docx4j.dml.STTextAnchoringType;
import org.docx4j.dml.STTextVertOverflowType;
import org.docx4j.dml.STTextVerticalType;
import org.docx4j.dml.STTextWrappingType;
import org.docx4j.dml.TextFont;
import org.docx4j.dml.chart.CTAxDataSource;
import org.docx4j.dml.chart.CTAxPos;
import org.docx4j.dml.chart.CTCatAx;
import org.docx4j.dml.chart.CTChart;
import org.docx4j.dml.chart.CTChartSpace;
import org.docx4j.dml.chart.CTCrossBetween;
import org.docx4j.dml.chart.CTCrosses;
import org.docx4j.dml.chart.CTDispBlanksAs;
import org.docx4j.dml.chart.CTLayout;
import org.docx4j.dml.chart.CTLblAlgn;
import org.docx4j.dml.chart.CTLblOffset;
import org.docx4j.dml.chart.CTLegend;
import org.docx4j.dml.chart.CTLegendPos;
import org.docx4j.dml.chart.CTMarker;
import org.docx4j.dml.chart.CTMarkerStyle;
import org.docx4j.dml.chart.CTNumData;
import org.docx4j.dml.chart.CTNumDataSource;
import org.docx4j.dml.chart.CTNumRef;
import org.docx4j.dml.chart.CTNumVal;
import org.docx4j.dml.chart.CTOrientation;
import org.docx4j.dml.chart.CTPlotArea;
import org.docx4j.dml.chart.CTRelId;
import org.docx4j.dml.chart.CTScaling;
import org.docx4j.dml.chart.CTSerTx;
import org.docx4j.dml.chart.CTStrData;
import org.docx4j.dml.chart.CTStrRef;
import org.docx4j.dml.chart.CTStrVal;
import org.docx4j.dml.chart.CTTextLanguageID;
import org.docx4j.dml.chart.CTTickLblPos;
import org.docx4j.dml.chart.CTTickMark;
import org.docx4j.dml.chart.CTTitle;
import org.docx4j.dml.chart.CTTx;
import org.docx4j.dml.chart.CTValAx;
import org.docx4j.dml.chart.STAxPos;
import org.docx4j.dml.chart.STCrossBetween;
import org.docx4j.dml.chart.STCrosses;
import org.docx4j.dml.chart.STDispBlanksAs;
import org.docx4j.dml.chart.STLblAlgn;
import org.docx4j.dml.chart.STLegendPos;
import org.docx4j.dml.chart.STMarkerStyle;
import org.docx4j.dml.chart.STOrientation;
import org.docx4j.dml.chart.STTickLblPos;
import org.docx4j.dml.chart.STTickMark;
import org.docx4j.dml.spreadsheetdrawing.CTAnchorClientData;
import org.docx4j.dml.spreadsheetdrawing.CTDrawing;
import org.docx4j.dml.spreadsheetdrawing.CTGraphicalObjectFrame;
import org.docx4j.dml.spreadsheetdrawing.CTGraphicalObjectFrameNonVisual;
import org.docx4j.dml.spreadsheetdrawing.CTTwoCellAnchor;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.exceptions.InvalidFormatException;
import org.docx4j.openpackaging.packages.SpreadsheetMLPackage;
import org.docx4j.openpackaging.parts.PartName;
import org.docx4j.openpackaging.parts.DrawingML.Chart;
import org.docx4j.openpackaging.parts.DrawingML.Drawing;
import org.docx4j.openpackaging.parts.SpreadsheetML.WorksheetPart;

import com.jumbletree.docx5j.xlsx.builders.BuilderMethods;
import com.jumbletree.docx5j.xlsx.builders.WorkbookBuilder;

public abstract class XLSXChart implements BuilderMethods {

	private WorkbookBuilder factory;

	public XLSXChart(WorksheetPart sheet, SpreadsheetMLPackage pkg, WorkbookBuilder factory) throws InvalidFormatException {
		String prefix = "/xl/charts/chart";
		this.chartNumber = getNextPartNumber(prefix, pkg);
		PartName chartPartName = new PartName(prefix + chartNumber + ".xml");

		this.chart = new Chart(chartPartName);
		this.sheet = sheet;
		this.pkg = pkg;
		this.factory = factory;
	}

	class SeriesTitle {
		boolean isReference;
		String value;
		public SeriesTitle(boolean isReference, String value) {
			super();
			this.isReference = isReference;
			this.value = value;
		}
		public boolean isReference() {
			return isReference;
		}
		public String getValue() {
			return value;
		}
	}

	protected int chartNumber;
	protected Chart chart;
	protected WorksheetPart sheet;
	protected SpreadsheetMLPackage pkg;

	//Properties of the chart
	Color chartFillColor = null;
	Color chartLineColor = new Color(0xB3, 0xB3, 0xB3);
	Color chartAreaFillColor = Color.white;
	Color chartAreaLineColor = null;
	Color legendFillColor = null;
	Color legendLineColor = null;
	Location legendPosition = Location.RIGHT;
	private String title;
	String xAxisLabel;
	String yAxisLabel;
	
	public void create(XLSXRange cellAnchor) throws Docx4JException {
		//Create the actual chart object's xml
		CTChartSpace chartSpace = new CTChartSpace();
		//Use 1900 dates
		chartSpace.setDate1904(createBoolean(false));
		chartSpace.setRoundedCorners(createBoolean(false));
		
		CTTextLanguageID lang = new CTTextLanguageID();
		lang.setVal("en-NZ");
		chartSpace.setLang(lang);

//		chartSpace.setRoundedCorners(createBoolean(false));

		//Create the chart
		CTChart chart = new CTChart();
		if (title == null) {
			chart.setAutoTitleDeleted(createBoolean(false));
		} 
		//This is the excel version
		CTTitle title = new CTTitle();
		title.setLayout(new CTLayout());
		title.setOverlay(createBoolean(false));
		
		CTShapeProperties sppr = new CTShapeProperties();
		sppr.setNoFill(new CTNoFillProperties());
		CTLineProperties ln = new CTLineProperties();
		ln.setNoFill(new CTNoFillProperties());
		sppr.setLn(ln);
		sppr.setEffectLst(new CTEffectList());
		
		
		CTTx tx = new CTTx();
		CTTextBody txpr = new CTTextBody();
		tx.setRich(txpr);
		
		CTTextBodyProperties bodyPr = new CTTextBodyProperties();
		bodyPr.setRot(0);
		bodyPr.setSpcFirstLastPara(false);
		bodyPr.setVertOverflow(STTextVertOverflowType.ELLIPSIS);
		bodyPr.setVert(STTextVerticalType.HORZ);
		bodyPr.setWrap(STTextWrappingType.SQUARE);
		bodyPr.setAnchor(STTextAnchoringType.CTR);
		bodyPr.setAnchorCtr(true);
		txpr.setBodyPr(bodyPr);
		
		txpr.setLstStyle(new CTTextListStyle());
		
		CTTextParagraph p = new CTTextParagraph();
		CTTextParagraphProperties ppr = new CTTextParagraphProperties();
		CTTextCharacterProperties defrpr = new CTTextCharacterProperties();
		ppr.setDefRPr(defrpr);
		p.setPPr(ppr);
		
		CTRegularTextRun run = new CTRegularTextRun();
		CTTextCharacterProperties rpr = new CTTextCharacterProperties();
		rpr.setSz(1400);
		CTSolidColorFillProperties fill = new CTSolidColorFillProperties();
		fill.setSrgbClr(createRGBColor(new Color(0x59, 0x59, 0x59)));
		rpr.setSolidFill(fill);
		TextFont font = new TextFont();
		font.setTypeface("Calibri");
		rpr.setLatin(font);
		run.setRPr(rpr);
		run.setT(this.title);
		p.getEGTextRun().add(run);
		txpr.getP().add(p);

		title.setTx(tx);
		title.setOverlay(createBoolean(false));
		chart.setTitle(title);
		chart.setAutoTitleDeleted(createBoolean(false));
		
		CTPlotArea ctp = new CTPlotArea();
		chart.setPlotArea(ctp);

		CTLayout layout = new CTLayout();
		ctp.setLayout(layout);

		createChart(chart, ctp);

		setupAreaProps(ctp);

		//Legend etc
		setupLegend(chart);

		chartSpace.setChart(chart);

		//spPr
		CTShapeProperties chartAreaProps = new CTShapeProperties();
		fill = new CTSolidColorFillProperties();
		CTSRgbColor color = new CTSRgbColor();
		color.setVal("ffffff");
		fill.setSrgbClr(color);
		chartAreaProps.setSolidFill(fill);
		chartSpace.setSpPr(chartAreaProps);
		this.chart.setJaxbElement(chartSpace);

		//Does a drawing already exist?
		CTDrawing drawing = null;
		Drawing drawingPart = null;
		org.xlsx4j.sml.CTDrawing drawingRel = sheet.getContents().getDrawing();
		if (drawingRel != null) {
			String id = drawingRel.getId();
			drawingPart = (Drawing)sheet.relationships.getPart(id);
			drawing = drawingPart.getContents();
		} else {
			String prefix = "/xl/drawings/drawing";
			int chartNumber = getNextPartNumber(prefix, pkg);

			PartName drawingPartName = new PartName(prefix + chartNumber + ".xml");
			drawingPart = new Drawing(drawingPartName);
			//Add it all in, so that we can get the ids out
			String drawingId = sheet.addTargetPart(drawingPart).getId();
			//And now (finally) add it to the sheet

			drawingRel = new org.xlsx4j.sml.CTDrawing();
			drawingRel.setId(drawingId);
			sheet.getContents().setDrawing(drawingRel);
			//Create the actual drawing object's xml
			drawing = new CTDrawing();
			drawingPart.setJaxbElement(drawing);
		}


		//Now add a drawing that the chart will be put into
		String chartId = drawingPart.addTargetPart(this.chart).getId();


		CTTwoCellAnchor anchor = new CTTwoCellAnchor();
		org.docx4j.dml.spreadsheetdrawing.CTMarker from = new org.docx4j.dml.spreadsheetdrawing.CTMarker();
		from.setCol(cellAnchor.startCellNumericColumn());
		from.setRow(cellAnchor.startCellNumericRow());      
		anchor.setFrom(from);
		org.docx4j.dml.spreadsheetdrawing.CTMarker to = new org.docx4j.dml.spreadsheetdrawing.CTMarker();
		to.setCol(cellAnchor.endCellNumericColumn());
		to.setRow(cellAnchor.endCellNumericRow());
		anchor.setTo(to);

		//the frame
		CTGraphicalObjectFrame frame = new CTGraphicalObjectFrame();
		frame.setMacro("");
		CTGraphicalObjectFrameNonVisual pr = new CTGraphicalObjectFrameNonVisual();
		CTNonVisualDrawingProps nv = new CTNonVisualDrawingProps();
		nv.setName("Chart " + chartNumber);
		nv.setId(2L);
		pr.setCNvPr(nv);
		pr.setCNvGraphicFramePr(new CTNonVisualGraphicFrameProperties());
		frame.setNvGraphicFramePr(pr);

		CTTransform2D trans = new CTTransform2D();
		CTPoint2D off = new CTPoint2D();
		off.setX(0);
		off.setY(0);
		trans.setOff(off);
		CTPositiveSize2D ext = new CTPositiveSize2D();
		ext.setCx(0);
		ext.setCy(0);
		trans.setExt(ext);
		frame.setXfrm(trans);

		//The graphic === the chart
		Graphic graphic = new Graphic();
		GraphicData data = new GraphicData();
		data.setUri("http://schemas.openxmlformats.org/drawingml/2006/chart");

		CTRelId rel = new CTRelId();
		rel.setId(chartId);
		data.getAny().add(new JAXBElement<CTRelId>(new QName("http://schemas.openxmlformats.org/drawingml/2006/chart", "chart", "c"), CTRelId.class, rel));

		graphic.setGraphicData(data);
		frame.setGraphic(graphic);
		anchor.setGraphicFrame(frame);
		anchor.setClientData(new CTAnchorClientData());
		drawing.getEGAnchor().add(anchor);

	}

	protected CTTitle createTitle(int size, String text, STTextAnchoringType anchor, boolean anchorCtr, Integer rotate) {
		CTTitle title = new CTTitle();
		CTTx tx = new CTTx();
		CTTextBody rich = new CTTextBody();
		CTTextBodyProperties bodyPr = new CTTextBodyProperties();
//		if (anchor != null) {
//			bodyPr.setAnchor(anchor);
//			bodyPr.setAnchorCtr(new Boolean(anchorCtr));
//		}
//		if (rotate != null) {
//			bodyPr.setRot(rotate * 60000);
//			bodyPr.setVert(STTextVerticalType.HORZ);
//		}
		rich.setBodyPr(bodyPr);
		rich.setLstStyle(new CTTextListStyle());
		CTTextParagraph p = new CTTextParagraph();
		CTTextParagraphProperties ppr = new CTTextParagraphProperties();
		ppr.setDefRPr(new CTTextCharacterProperties());
		p.setPPr(ppr);
		
		CTRegularTextRun r = new CTRegularTextRun();
		CTTextCharacterProperties rpr = new CTTextCharacterProperties();
//		rpr.setSz(size);
		TextFont font = new TextFont();
		font.setTypeface("Calibri");
		rpr.setLatin(font);
		r.setRPr(rpr);
		r.setT(text);
		p.getEGTextRun().add(r);
		rich.getP().add(p);
		tx.setRich(rich);
		title.setTx(tx);
		title.setLayout(new CTLayout());
		title.setOverlay(createBoolean(false));

		return title;
	}

	protected abstract void createChart(CTChart chart2, CTPlotArea ctp) throws Docx4JException;

	protected void setupAreaProps(CTPlotArea ctp) {
		CTShapeProperties plotAreaProps = new CTShapeProperties();
		plotAreaProps.setNoFill(new CTNoFillProperties());
		CTLineProperties ln = new CTLineProperties();
		CTSolidColorFillProperties fill = new CTSolidColorFillProperties();
		CTSRgbColor color = new CTSRgbColor();
		color.setVal("b3b3b3");
		fill.setSrgbClr(color);
		ln.setSolidFill(fill);
		plotAreaProps.setLn(ln);
		ctp.setSpPr(plotAreaProps);
	}

	protected void setupLegend(CTChart chart) {
		CTLegend legend = new CTLegend();
		CTLegendPos pos = new CTLegendPos();
		switch (legendPosition) {
			case TOP: pos.setVal(STLegendPos.T); break;
			case LEFT: pos.setVal(STLegendPos.L); break;
			case BOTTOM: pos.setVal(STLegendPos.B); break;
			case RIGHT: pos.setVal(STLegendPos.R); break;
		}
		legend.setLegendPos(pos);
		legend.setLayout(new CTLayout());
		legend.setOverlay(createBoolean(false));
		chart.setLegend(legend);
		chart.setPlotVisOnly(createBoolean(true));
		CTDispBlanksAs as = new CTDispBlanksAs();
		as.setVal(STDispBlanksAs.GAP);
		chart.setDispBlanksAs(as);
		chart.setShowDLblsOverMax(createBoolean(false));
	}

	protected void setupAxes(CTPlotArea ctp, int cataxId, int valaxId) throws Docx4JException {
		//Axes
		CTCatAx catax = new CTCatAx();
		catax.setAxId(createUnsignedInt(cataxId));
		CTScaling scaling = new CTScaling();
		CTOrientation orient = createVal(CTOrientation.class, STOrientation.MIN_MAX);
		orient.setVal(STOrientation.MIN_MAX);
		scaling.setOrientation(orient);
		catax.setScaling(scaling);
		catax.setDelete(createBoolean(false));
		CTAxPos axpos = new CTAxPos();
		axpos.setVal(STAxPos.B);
		catax.setAxPos(axpos);
		catax.setMajorTickMark(createVal(CTTickMark.class, STTickMark.OUT));
		catax.setMinorTickMark(createVal(CTTickMark.class, STTickMark.NONE));
		catax.setTickLblPos(createVal(CTTickLblPos.class, STTickLblPos.NEXT_TO));
		CTShapeProperties sppr = new CTShapeProperties();
		CTLineProperties ln = new CTLineProperties();
		ln.setW(9360);
		sppr.setLn(ln);
		CTSolidColorFillProperties fill = new CTSolidColorFillProperties();
		fill.setSrgbClr(createRGBColor(new Color(0xb9, 0xb9, 0xb9)));
		ln.setSolidFill(fill);
		ln.setRound(new CTLineJoinRound());
		catax.setSpPr(sppr);
		catax.setCrossAx(createUnsignedInt(valaxId));
		catax.setCrosses(createVal(CTCrosses.class, STCrosses.AUTO_ZERO));
		catax.setAuto(createBoolean(true));
		catax.setLblAlgn(createVal(CTLblAlgn.class, STLblAlgn.CTR));
		CTLblOffset offset = new CTLblOffset();
		offset.setVal(100);
		catax.setLblOffset(offset);
		catax.setNoMultiLvlLbl(createBoolean(true));
		if (getXAxisLabel() != null) {
			catax.setTitle(createTitle(900, getXAxisLabel(), null, false, null));
		}

		CTValAx valax = new CTValAx();
		valax.setAxId(createUnsignedInt(valaxId));
		scaling = new CTScaling();
		CTOrientation orientation = new CTOrientation();
		orientation.setVal(STOrientation.MIN_MAX);
		scaling.setOrientation(orientation);
		valax.setScaling(scaling);
		valax.setDelete(createBoolean(false));
		CTAxPos pos = new CTAxPos();
		pos.setVal(STAxPos.L);
		valax.setAxPos(pos);
		if (getYAxisLabel() != null) {
			valax.setTitle(createTitle(900, getYAxisLabel(), null, false, -90));
		}
		CTTickMark tick = new CTTickMark();
		tick.setVal(STTickMark.OUT);
		valax.setMajorTickMark(tick);
		tick = new CTTickMark();
		tick.setVal(STTickMark.NONE);
		valax.setMinorTickMark(tick);
		CTTickLblPos lblpos = new CTTickLblPos();
		lblpos.setVal(STTickLblPos.NEXT_TO);
		valax.setTickLblPos(lblpos);
		sppr = new CTShapeProperties();
		ln = new CTLineProperties();
		ln.setW(9360);
		sppr.setLn(ln);
		fill = new CTSolidColorFillProperties();
		fill.setSrgbClr(createRGBColor(new Color(0xb9, 0xb9, 0xb9)));
		ln.setSolidFill(fill);
		valax.setSpPr(sppr);
		valax.setCrossAx(createUnsignedInt(cataxId));
		valax.setCrossesAt(createDouble(1));
		CTCrossBetween between = new CTCrossBetween();
		between.setVal(STCrossBetween.MID_CAT);
		valax.setCrossBetween(between);
		ctp.getValAxOrCatAxOrDateAx().add(catax);
		ctp.getValAxOrCatAxOrDateAx().add(valax);
	}

	protected CTSerTx createSeriesTitle(SeriesTitle title) throws Docx4JException {
		CTSerTx tx = new CTSerTx();
		if (title.isReference()) {
			CTStrRef ref = createStrRef(XLSXRange.fromReference(title.getValue()));
			tx.setStrRef(ref);
		} else {
			tx.setV(title.getValue());
		}
		return tx;
	}

	protected CTNumDataSource createNumDataSource(XLSXRange data) throws Docx4JException {
		CTNumDataSource dataSource = new CTNumDataSource();
		CTNumRef ref = createNumRef(data);
		dataSource.setNumRef(ref);
		return dataSource;
	}
	protected CTAxDataSource createAxDataSource(XLSXRange data, boolean isNum) throws Docx4JException {
		CTAxDataSource dataSource = new CTAxDataSource();
		if (isNum) {
			CTNumRef ref = createNumRef(data);
			dataSource.setNumRef(ref);
		} else {
			CTStrRef ref = createStrRef(data);
			dataSource.setStrRef(ref);
		}
		return dataSource;
	}

	protected CTNumRef createNumRef(XLSXRange data) throws Docx4JException {
		CTNumRef ref = new CTNumRef();
		
		ref.setF(data.absoluteReference());
		//Cache
		CTNumData cache = new CTNumData();
		cache.setFormatCode("General");
		cache.setPtCount(createUnsignedInt(data.getLinearSize()));
		int index = 0;
		for (XLSXRange cell : data.getLinearCells()) {
			CTNumVal val = new CTNumVal();
			val.setIdx(index++);
			val.setV(factory.getCellValueString(cell));
			cache.getPt().add(val);
		}
//		ref.setNumCache(cache);
		return ref;
	}

	protected CTStrRef createStrRef(XLSXRange data) throws Docx4JException {
		CTStrRef ref = new CTStrRef();
		ref.setF(data.absoluteReference());
		//Create the cache also
//		ref.setStrCache(createStringCache(data));
		return ref;
	}

	protected CTStrData createStringCache(XLSXRange data) throws Docx4JException {
		CTStrData cache = new CTStrData();
		cache.setPtCount(createUnsignedInt(data.getLinearSize()));
		int index = 0;
		for (XLSXRange cell : data.getLinearCells()) {
			CTStrVal pt = new CTStrVal();
			pt.setIdx(index++);
			pt.setV(factory.getCellValueString(cell));
			cache.getPt().add(pt);
		}
		return cache;
	}

	protected CTMarker createMarkerProperties(MarkerProperties markerProps) {
		//TODO define the marker properly in the series defs
		//Marker
		CTMarker marker = new CTMarker();
		CTMarkerStyle style = new CTMarkerStyle();
		
		if (markerProps == null) {
			style.setVal(STMarkerStyle.NONE);
		}
		marker.setSymbol(style);
		return marker;
	}

	protected CTShapeProperties createLineProperties(LineProperties line) {
		if (line != null) {
			CTShapeProperties sppr = new CTShapeProperties();
			CTSolidColorFillProperties fill = new CTSolidColorFillProperties();
//			fill.setSrgbClr(createColor(line.getColor()));
//			sppr.setSolidFill(fill);
			CTLineProperties ln = new CTLineProperties();
			ln.setW(line.getWidth());
//			fill = new CTSolidColorFillProperties();
			fill.setSrgbClr(createRGBColor(line.getColor()));
			ln.setSolidFill(fill);
			//Configure via properties
			ln.setRound(new CTLineJoinRound());
			sppr.setLn(ln);
			return sppr;
		}
		return null;
	}
	public Color getChartFillColor() {
		return chartFillColor;
	}

	public void setChartFillColor(Color chartFillColor) {
		this.chartFillColor = chartFillColor;
	}

	public Color getChartLineColor() {
		return chartLineColor;
	}

	public void setChartLineColor(Color chartLineColor) {
		this.chartLineColor = chartLineColor;
	}

	public Color getChartAreaFillColor() {
		return chartAreaFillColor;
	}

	public void setChartAreaFillColor(Color chartAreaFillColor) {
		this.chartAreaFillColor = chartAreaFillColor;
	}

	public Color getChartAreaLineColor() {
		return chartAreaLineColor;
	}

	public void setChartAreaLineColor(Color chartAreaLineColor) {
		this.chartAreaLineColor = chartAreaLineColor;
	}

	public Color getLegendFillColor() {
		return legendFillColor;
	}

	public void setLegendFillColor(Color legendFillColor) {
		this.legendFillColor = legendFillColor;
	}

	public Color getLegendLineColor() {
		return legendLineColor;
	}

	public void setLegendLineColor(Color legendLineColor) {
		this.legendLineColor = legendLineColor;
	}

	public Location getLegendPosition() {
		return legendPosition;
	}

	public void setLegendPosition(Location legendPosition) {
		this.legendPosition = legendPosition;
	}

	public String getTitle() {
		return title;
	}

	public void setTitle(String title) {
		this.title = title;
	}

	public String getXAxisLabel() {
		return xAxisLabel;
	}

	public void setXAxisLabel(String xAxisLabel) {
		this.xAxisLabel = xAxisLabel;
	}

	public String getYAxisLabel() {
		return yAxisLabel;
	}

	public void setYAxisLabel(String yAxisLabel) {
		this.yAxisLabel = yAxisLabel;
	}

}
