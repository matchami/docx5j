package com.jumbletree.docx5j.xlsx.builders;

import java.awt.Color;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.io.StringReader;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Set;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import javax.xml.bind.JAXBContext;
import javax.xml.bind.JAXBElement;
import javax.xml.bind.JAXBException;
import javax.xml.namespace.QName;
import javax.xml.transform.stream.StreamSource;

import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.exceptions.InvalidFormatException;
import org.docx4j.openpackaging.io3.Save;
import org.docx4j.openpackaging.packages.SpreadsheetMLPackage;
import org.docx4j.openpackaging.parts.PartName;
import org.docx4j.openpackaging.parts.VMLPart;
import org.docx4j.openpackaging.parts.SpreadsheetML.CommentsPart;
import org.docx4j.openpackaging.parts.SpreadsheetML.PrinterSettings;
import org.docx4j.openpackaging.parts.SpreadsheetML.SharedStrings;
import org.docx4j.openpackaging.parts.SpreadsheetML.Styles;
import org.docx4j.openpackaging.parts.SpreadsheetML.WorksheetPart;
import org.docx4j.openpackaging.parts.relationships.RelationshipsPart;
import org.docx4j.relationships.Relationship;
import org.docx4j.sharedtypes.STVerticalAlignRun;
import org.docx4j.vml.CTPath;
import org.docx4j.vml.CTShadow;
import org.docx4j.vml.CTShape;
import org.docx4j.vml.CTShapetype;
import org.docx4j.vml.CTStroke;
import org.docx4j.vml.CTTextbox;
import org.docx4j.vml.STExt;
import org.docx4j.vml.STStrokeJoinStyle;
import org.docx4j.vml.STTrueFalse;
import org.docx4j.vml.officedrawing.CTIdMap;
import org.docx4j.vml.officedrawing.CTShapeLayout;
import org.docx4j.vml.officedrawing.STConnectType;
import org.docx4j.vml.officedrawing.STInsetMode;
import org.docx4j.vml.root.Xml;
import org.docx4j.vml.spreadsheetDrawing.CTClientData;
import org.xlsx4j.jaxb.Context;
import org.xlsx4j.sml.CTAuthors;
import org.xlsx4j.sml.CTBooleanProperty;
import org.xlsx4j.sml.CTBorder;
import org.xlsx4j.sml.CTBorderPr;
import org.xlsx4j.sml.CTBorders;
import org.xlsx4j.sml.CTCellAlignment;
import org.xlsx4j.sml.CTCellProtection;
import org.xlsx4j.sml.CTCellFormula;
import org.xlsx4j.sml.CTCellStyle;
import org.xlsx4j.sml.CTCellStyleXfs;
import org.xlsx4j.sml.CTCellStyles;
import org.xlsx4j.sml.CTCellXfs;
import org.xlsx4j.sml.CTColor;
import org.xlsx4j.sml.CTComment;
import org.xlsx4j.sml.CTCommentList;
import org.xlsx4j.sml.CTComments;
import org.xlsx4j.sml.CTFill;
import org.xlsx4j.sml.CTFills;
import org.xlsx4j.sml.CTFont;
import org.xlsx4j.sml.CTFontFamily;
import org.xlsx4j.sml.CTFontName;
import org.xlsx4j.sml.CTFontScheme;
import org.xlsx4j.sml.CTFontSize;
import org.xlsx4j.sml.CTFonts;
import org.xlsx4j.sml.CTHyperlink;
import org.xlsx4j.sml.CTHyperlinks;
import org.xlsx4j.sml.CTLegacyDrawing;
import org.xlsx4j.sml.CTNumFmt;
import org.xlsx4j.sml.CTNumFmts;
import org.xlsx4j.sml.CTPageSetup;
import org.xlsx4j.sml.CTPatternFill;
import org.xlsx4j.sml.CTRElt;
import org.xlsx4j.sml.CTRPrElt;
import org.xlsx4j.sml.CTRst;
import org.xlsx4j.sml.CTSst;
import org.xlsx4j.sml.CTStylesheet;
import org.xlsx4j.sml.CTUnderlineProperty;
import org.xlsx4j.sml.CTVerticalAlignFontProperty;
import org.xlsx4j.sml.CTXf;
import org.xlsx4j.sml.CTXstringWhitespace;
import org.xlsx4j.sml.Cell;
import org.xlsx4j.sml.ObjectFactory;
import org.xlsx4j.sml.Row;
import org.xlsx4j.sml.STBorderStyle;
import org.xlsx4j.sml.STCellType;
import org.xlsx4j.sml.STFontScheme;
import org.xlsx4j.sml.STOrientation;
import org.xlsx4j.sml.STPatternType;
import org.xlsx4j.sml.STUnderlineValues;
import org.xlsx4j.sml.Sheet;
import org.xlsx4j.sml.SheetData;

import com.jumbletree.docx5j.xlsx.CommentPosition;
import com.jumbletree.docx5j.xlsx.LineChart;
import com.jumbletree.docx5j.xlsx.ScatterChart;
import com.jumbletree.docx5j.xlsx.XLSXRange;

public class WorkbookBuilder implements BuilderMethods {

	public class ObjectKey {

		private Object[] data;
		
		public ObjectKey(Object ... data) {
			this.data = data;
		}

		@Override
		public int hashCode() {
			return Arrays.hashCode(data);
		}

		@Override
		public boolean equals(Object obj) {
			if (obj == this)
				return true;
			if (!(obj instanceof ObjectKey))
				return false;
			
			return Arrays.equals(data, ((ObjectKey)obj).data);
		}

		
	}

	private SpreadsheetMLPackage pkg;
	private ArrayList<WorksheetPart> sheets;
	private CTStylesheet stylesheet;
	private ObjectFactory factory;
	private CTSst strings;
	private HashMap<String, Integer> stringCache;
	private HashMap<String, Long> styles = new HashMap<>();
	private HashMap<Long, Long> unlockedStyles = new HashMap<>();
	private HashMap<ObjectKey, Long> objectCache = new HashMap<>();
	private List<CommentsPart> comments = new ArrayList<>();
	private List<VMLPart> commentsDrawings = new ArrayList<>();
	private Set<Long> thickBottomStyles = new HashSet<>();
	private int userDefinedFormatIndex = 164;
	private HashMap<String, Long> userDefinedFormats = new HashMap<>();

	public WorkbookBuilder() throws InvalidFormatException, JAXBException {
		pkg = SpreadsheetMLPackage.createPackage();
		factory = new ObjectFactory();
		
		//Create the default sheet
		WorksheetPart sheet = pkg.createWorksheetPart(new PartName("/xl/worksheets/sheet1.xml"), "Sheet1", 1);
		sheets = new ArrayList<WorksheetPart>();
		sheets.add(sheet);

		//Create the default styles
		Styles styles = new Styles(new PartName("/xl/styles.xml"));
		this.stylesheet = new CTStylesheet();
		styles.setJaxbElement(stylesheet);
		pkg.getWorkbookPart().addTargetPart(styles);

		//Add a bunch more defaults
		createFont("Calibri", 11, Color.black, false, false, null);
		createDefaultFill();
		createDefaultBorder();
		createDefaultCellStyleXf();
		createDefaultCellXf();
		createDefaultCellStyle();
		
		//SharedStrings
		SharedStrings strings = new SharedStrings(new PartName("/xl/sharedStrings.xml"));
		strings.setJaxbElement(this.strings = new CTSst());
		pkg.getWorkbookPart().addTargetPart(strings);
		this.stringCache = new HashMap<String, Integer>();
	}
	
	public void installStyle(String name, int index) {
		this.styles.put(name, new Long(index));
	}

	public WorksheetBuilder getSheet(int index) {
		if (index >= sheets.size()) {
			throw new IllegalArgumentException("Workbook does not have " + (index+1) + " sheets");
		}
		return new WorksheetBuilder(index, sheets.get(index), this);
	}

	public WorksheetBuilder appendSheet() throws Docx4JException, JAXBException {
		int id = sheets.size() + 1;
		//Work out the sheet filename
		String filename = "/xl/worksheets/sheet" + id + ".xml";
		
		//Now work out the sheet name ... follow same logic as Excel
		int max = 0;
		
		Pattern pattern = Pattern.compile("^Sheet(\\d+)$");
		for (Sheet sheet : pkg.getWorkbookPart().getContents().getSheets().getSheet()) {
			Matcher matcher = pattern.matcher(sheet.getName());
			if (matcher.matches()) {
				max = Math.max(max, Integer.parseInt(matcher.group(1)));
			}
		}
		
		String name = "Sheet" + (max+1);
		
		WorksheetPart sheet = pkg.createWorksheetPart(new PartName(filename), name, id);
		sheets.add(sheet);
		
		return new WorksheetBuilder(id-1, sheet, this);
	}

	private void createDefaultCellStyle() {
		stylesheet.setCellStyles(new CTCellStyles());
		CTCellStyle style = new CTCellStyle();
		style.setName("Normal");
		style.setXfId(0L);
		style.setBuiltinId(0L);
		stylesheet.getCellStyles().getCellStyle().add(style);
	}

	private void createDefaultCellStyleXf() {
		stylesheet.setCellStyleXfs(new CTCellStyleXfs());
		CTXf xf = new CTXf();
		xf.setNumFmtId(0L);
		xf.setFontId(0L);
		xf.setFillId(0L);
		xf.setBorderId(0L);
		stylesheet.getCellStyleXfs().getXf().add(xf);
	}

	private void createDefaultCellXf() {
		stylesheet.setCellXfs(new CTCellXfs());
		CTXf xf = new CTXf();
		xf.setNumFmtId(0L);
		xf.setFontId(0L);
		xf.setFillId(0L);
		xf.setBorderId(0L);
		xf.setXfId(0L);
		stylesheet.getCellXfs().getXf().add(xf);
	}

	protected void createDefaultBorder() {
		CTBorder border = new CTBorder();
		border.setBottom(new CTBorderPr());
		border.setLeft(new CTBorderPr());
		border.setRight(new CTBorderPr());
		border.setTop(new CTBorderPr());
		border.setDiagonal(new CTBorderPr());
		stylesheet.setBorders(new CTBorders());
		stylesheet.getBorders().getBorder().add(border);
	}

	/**
	 * These two fills need to be in place, or excel has issues
	 */
	protected void createDefaultFill() {
		CTFills fills = new CTFills();
		stylesheet.setFills(fills);
		createFill(null, null, STPatternType.NONE);
		createFill(null, null, STPatternType.GRAY_125);
	}

	int createStyle(String formatDefinition, Long fontId, Long fillId, Long borderId, CTCellAlignment alignment) {
		Long formatId = this.userDefinedFormats.get(formatDefinition);
		if (formatId == null) {
			formatId = Long.valueOf(userDefinedFormatIndex++);
			this.userDefinedFormats.put(formatDefinition, formatId);
			//And add it to the definitions
			CTNumFmts formats = stylesheet.getNumFmts();
			if (formats == null) {
				formats = new CTNumFmts();
				stylesheet.setNumFmts(formats);
			}
			CTNumFmt format = new CTNumFmt();
			format.setNumFmtId(formatId.longValue());
			format.setFormatCode(formatDefinition);
			formats.getNumFmt().add(format);
		}
		return createStyle(formatId, fontId, fillId, borderId, alignment);
	}

	int createStyle(Long formatId, Long fontId, Long fillId, Long borderId, CTCellAlignment alignment) {
		CTXf xf = new CTXf();
		xf.setNumFmtId(formatId == null ? 0 : formatId);
		xf.setFontId(fontId == null ? 0L : fontId);
		xf.setFillId(fillId == null ? 0L : fillId);
		xf.setBorderId(borderId == null ? 0L : borderId);
		if (formatId != null) {
			xf.setApplyNumberFormat(true);
		}
		if (fontId != null) {
			xf.setApplyFont(true);
		}
		if (fillId != null) {
			xf.setApplyFill(true);
		}
		if (borderId != null) {
			xf.setApplyBorder(true);
		}
		//Always ref the default style
		xf.setXfId(0L);
		
		xf.setAlignment(alignment);
		
		int index = stylesheet.getCellXfs().getXf().size();
		stylesheet.getCellXfs().getXf().add(xf);
		
		return index;
	}

	CTXf getStyle(int index) {
		return stylesheet.getCellXfs().getXf().get(index);
	}
	
	CTFont getStyleFont(int index) {
		CTXf style = getStyle(index);
		
		Long font = style.getFontId();
		if (font == null) {
			throw new IllegalArgumentException("Cannot retrieve font from unstyled cell");
		}
		
		return getFont(font.intValue());
	}
	
	CTFont getFont(int index) {
		return stylesheet.getFonts().getFont().get(index);
	}
	
	protected CTXf createXf(Long formatId, Long fontId, Long borderId) {
		CTXf xf = new CTXf();
		xf.setNumFmtId(formatId == null ? 0L : formatId);
		xf.setBorderId(borderId == null ? 0L : borderId);
		xf.setFontId(fontId == null ? 0L : fontId);
		
		xf.setApplyBorder(borderId != null);
		xf.setApplyFont(fontId != null);
		xf.setApplyNumberFormat(formatId != null);
		xf.setApplyAlignment(false);
		xf.setApplyProtection(false);
		return xf;
	}

	public long getDefaultXFIndex() {
		return 0;
	}
	
	public long getDefaultFontIndex() {
		return 0;
	}
	
	public long getGeneralFormatIndex() {
		return 164;
	}
	
	public long getEmptyBorderIndex() {
		return 0;
	}
	
	Long createBorder(STBorderStyle style, Color color) {
		return createBorder(style, color, style, color, style, color, style, color);
	}
	Long createBorder(STBorderStyle topStyle, Color topColor, STBorderStyle rightStyle, Color rightColor, STBorderStyle bottomStyle, Color bottomColor, STBorderStyle leftStyle, Color leftColor) {
		ObjectKey key = new ObjectKey(topStyle, topColor, rightStyle, rightColor, bottomStyle, bottomColor, leftStyle, leftColor);
		Long id = objectCache.get(key);
		if (id == null) {
			CTBorder border = new CTBorder();
			border.setBottom(createBorderPr(bottomStyle, bottomColor));
			border.setLeft(createBorderPr(leftStyle, leftColor));
			border.setRight(createBorderPr(rightStyle, rightColor));
			border.setTop(createBorderPr(topStyle, topColor));
			border.setDiagonal(new CTBorderPr());
	
			int index = stylesheet.getBorders().getBorder().size();
			stylesheet.getBorders().getBorder().add(border);

			objectCache.put(key, id = new Long(index));
		}
		return id;
	}

	Long createFill(Color bgColor, Color fgColor, STPatternType patternType) {
		ObjectKey key = new ObjectKey(bgColor, fgColor, patternType);
		Long id = objectCache.get(key);
		if (id == null) {

			CTFill fill = new CTFill();
			CTPatternFill pattern = new CTPatternFill();
			fill.setPatternFill(pattern);
		
			if (bgColor != null)
				pattern.setBgColor(createColor(bgColor));
			if (fgColor != null)
				pattern.setFgColor(createColor(fgColor));
			pattern.setPatternType(patternType);
			
			int index = stylesheet.getFills().getFill().size();
			stylesheet.getFills().getFill().add(fill);
			
			objectCache.put(key, id = new Long(index));
		}
		return id;
	}

	private CTBorderPr createBorderPr(STBorderStyle style, Color color) {
		CTBorderPr pr = new CTBorderPr();
		CTColor ctcolor = new CTColor();
		ctcolor.setRgb(getColorBytes(color));
		pr.setColor(ctcolor);
		pr.setStyle(style);
		return pr;
	}

	Long createFont(String fontName, int size, Color color, boolean bold, boolean italic, STUnderlineValues underline) {
		ObjectKey key = new ObjectKey(fontName, size, color, bold, italic, underline);
		Long id = objectCache.get(key);
		
		if (id == null) {
			CTFont font = new CTFont();
			List<JAXBElement<?>> fontProperties = font.getNameOrCharsetOrFamily();
			
			applyFont(fontName, size, color, bold, italic, underline, fontProperties);
			
			if (stylesheet.getFonts() == null) {
				stylesheet.setFonts(new CTFonts());
			}
			int index = stylesheet.getFonts().getFont().size();
			stylesheet.getFonts().getFont().add(font);
			
			objectCache.put(key, id = new Long(index));
		}
		return id;
	}

	protected void applyFont(String fontName, int size, Color color, boolean bold, boolean italic, STUnderlineValues underline, List<JAXBElement<?>> fontProperties) {
		setFontSize(size, fontProperties);
		setFontName(fontName, fontProperties);
		setFontColor(color, fontProperties);
		if (bold) {
			CTBooleanProperty bool = new CTBooleanProperty();
			bool.setVal(true);
			fontProperties.add(factory.createCTFontB(bool));
		}
		if (italic) {
			CTBooleanProperty bool = new CTBooleanProperty();
			bool.setVal(true);
			fontProperties.add(factory.createCTFontI(bool));
		}
		if (underline != null) {
			CTUnderlineProperty uline = new CTUnderlineProperty();
			uline.setVal(underline);
			fontProperties.add(factory.createCTFontU(uline));
		}
		CTFontFamily family = new CTFontFamily();
		family.setVal(2);
		fontProperties.add(factory.createCTFontFamily(family));
		//TODO figure out what the real requirements are here.
		if (fontName.equals("Calibri")) {
			CTFontScheme scheme = new CTFontScheme();
			scheme.setVal(STFontScheme.MINOR);
			fontProperties.add(factory.createCTFontScheme(scheme));
		}
	}
	
	private void setFontSize(long size, List<JAXBElement<?>> font){
		CTFontSize fontSize = new CTFontSize();
		fontSize.setVal(size);
		fontSize.setParent(font);
		font.add(factory.createCTFontSz(fontSize));
	}

	private void setFontColor(Color color, List<JAXBElement<?>> font){
		CTColor fontCol = new CTColor();
		fontCol.setRgb(getColorBytes(color));
		//fontCol.setTheme( new Long(1) );
		//fontCol.setTint( new Double(0.0) );
		fontCol.setParent(font);
		JAXBElement<CTColor> element1 = factory.createCTFontColor(fontCol);
		font.add(element1);
	}

	private void setFontName(String name, List<JAXBElement<?>> font){
		CTFontName fontName = new CTFontName();
		fontName.setVal(name);
		fontName.setParent(font);
		font.add(factory.createCTFontName(fontName));
	}

	public void setValue(int sheet, int col, int row, double value) throws Docx4JException {
		setValue(sheet, col, row, value, (Long)null);
	}
	
	public void setValue(int sheet, int col, int row, double value, String styleName) throws Docx4JException {
		setValue(sheet, col, row, value, styles.get(styleName));
	}
	
	public void setValue(int sheet, int col, int row, double value, Long style) throws Docx4JException {
		Cell theCell = getCell(sheet, col, row);

		theCell.setV(String.valueOf(value));
		if (style != null)
			theCell.setS(style);
	}

	public void setValue(int sheet, int col, int row, String value) throws Docx4JException {
		setValue(sheet, col, row, value, (Long)null);
	}

	public void setValue(int sheet, int col, int row, String value, String styleName) throws Docx4JException {
		setValue(sheet, col, row, value, styles.get(styleName));
	}
	
	public void setValue(int sheet, int col, int row, String value, Long style) throws Docx4JException {
		Cell theCell = getCell(sheet, col, row);

		theCell.setT(STCellType.S);

		int index = getStringCache(value);
		theCell.setV(String.valueOf(index));
		if (style != null)
			theCell.setS(style);
	}

	public void setFormula(int sheet, int col, int row, String formula, String styleName) throws Docx4JException {
		setFormula(sheet, col, row, formula, styles.get(styleName));
	}

	public void setFormula(int sheet, int col, int row, String formula, Long style) throws Docx4JException {
		Cell theCell = getCell(sheet, col, row);

		theCell.setT(STCellType.N);
		
		CTCellFormula f = Context.getsmlObjectFactory().createCTCellFormula();
		f.setT(org.xlsx4j.sml.STCellFormulaType.NORMAL);
		f.setValue(formula);
		theCell.setF(f);
		if (style != null)
			theCell.setS(style);
	}

	protected Cell getCell(int sheet, int col, int row) throws Docx4JException {
		SheetData sheetData = sheets.get(sheet).getContents().getSheetData();

		//Find the row
		int rowRef = row + 1;
		Row theRow = null;
		for (int i=0; i<sheetData.getRow().size(); i++) {
			Row aRow = sheetData.getRow().get(i);
			if (aRow.getR() == rowRef) {
				theRow = aRow;
				break;
			} else if (aRow.getR() > rowRef) {
				//Insert it
				theRow = Context.getsmlObjectFactory().createRow();
				theRow.setR((long)rowRef);
				sheetData.getRow().add(i, theRow);
				break;
			}
		}
		if (theRow == null) {
			theRow = Context.getsmlObjectFactory().createRow();
			theRow.setR((long)rowRef);
			sheetData.getRow().add(theRow);
		}
		String ref = XLSXRange.asCell(col, row);
		for (Cell cell : theRow.getC()) {
			if (cell.getR().equals(ref))
				return cell;
		}
		Cell cell = Context.getsmlObjectFactory().createCell();
		theRow.getC().add(cell);
		cell.setR(ref);
		return cell;
	}

	/**
	 * Creates aline chart on the last sheet in the workbook
	 * @throws Docx4JException 
	 */
	public LineChart createLineChart() throws Docx4JException {
		return new LineChart(sheets.get(sheets.size()-1), pkg, this);
	}

	/**
	 * 
	 * @param series a list of name, range pairs that will make up the chart eg createChart("First series", "Sheet1!$B$1:$B$20");
	 * @throws Docx4JException 
	 */
	public ScatterChart createScatterChart() throws Docx4JException {
		return new ScatterChart(sheets.get(sheets.size()-1), pkg, this);
	}

	public void save(OutputStream out) throws IOException, Docx4JException {
		//Finalise the styles
		if (stylesheet.getNumFmts() != null)
			stylesheet.getNumFmts().setCount((long)stylesheet.getNumFmts().getNumFmt().size());
		stylesheet.getFonts().setCount((long)stylesheet.getFonts().getFont().size());
		stylesheet.getFills().setCount((long)stylesheet.getFills().getFill().size());
		stylesheet.getBorders().setCount((long)stylesheet.getBorders().getBorder().size());
		stylesheet.getCellStyleXfs().setCount((long)stylesheet.getCellStyleXfs().getXf().size());
		stylesheet.getCellXfs().setCount((long)stylesheet.getCellXfs().getXf().size());
		stylesheet.getCellStyles().setCount((long)stylesheet.getCellStyles().getCellStyle().size());
		long count = strings.getSi().size();
		strings.setCount(count);
		strings.setUniqueCount(count);
		Save saver = new Save(pkg);
		saver.save(out);
		out.flush();
	}

	public String getCellValueString(XLSXRange cell) throws Docx4JException {
		long rowNum = cell.startCellNumericRow() + 1;	//R is user-speak not machine speak
		for (Sheet sheet : pkg.getWorkbookPart().getContents().getSheets().getSheet()) {
			if (sheet.getName().equals(cell.getSheet())) {
				WorksheetPart sheetPart = this.sheets.get((int)sheet.getSheetId()-1);
				for (Row row : sheetPart.getContents().getSheetData().getRow()) {
					if (row.getR().equals(rowNum)) {
						for (Cell c : row.getC()) {
							if (c.getR().equals(cell.getStartCell())) {
								switch (c.getT()) {
									case S:
										//Shared string
										int index = Integer.parseInt(c.getV());
										return strings.getSi().get(index).getT().getValue();
									case INLINE_STR:
									case N:
									default:
										return c.getV();
								}
							}
						}
					}
				}
			}
		}
		return null;
	}

	/**
	 * Sets the sheet's name.  Called by the sheet's factory, rather than by the user directly.  Use 
	 * getSheet(index).setName(name) instead.
	 */
	void setSheetName(int index, String name) throws Docx4JException {
		pkg.getWorkbookPart().getContents().getSheets().getSheet().get(index).setName(name);
	}

	public String getSheetName(int index) throws Docx4JException {
		return pkg.getWorkbookPart().getContents().getSheets().getSheet().get(index).getName();
	}

	public Long getStyle(String styleName) {
		return styles.get(styleName);
	}

	public int getStringCache(String value) {
		Integer index = stringCache.get(value);
		if (index == null) {
			index = strings.getSi().size();
			CTRst si = new CTRst();
			CTXstringWhitespace svalue = new CTXstringWhitespace();
			svalue.setValue(value);
			si.setT(svalue);
			strings.getSi().add(si);
			stringCache.put(value, index);
		}
		return index;
	}


	public void createHyperlink(WorksheetBuilder sheet, Cell cell, String url) throws Docx4JException {
		CTHyperlinks hyperlinks = sheet.sheet.getContents().getHyperlinks();
		if (hyperlinks == null) {
			hyperlinks = new CTHyperlinks();
			sheet.sheet.getContents().setHyperlinks(hyperlinks);
		}
		List<CTHyperlink> links = hyperlinks.getHyperlink();
		CTHyperlink link = new CTHyperlink();
		link.setDisplay("url");
		link.setRef(XLSXRange.fromCell(cell).singleCellSheetlessReference());
		
		//Get the rel
		RelationshipsPart rels = sheet.sheet.getRelationshipsPart(true);
		Relationship rel = new Relationship();
		String id = "rId" + rels.size() + 1;
		rel.setId(id);
		rel.setType("http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink");
		rel.setTarget(url);
		rel.setTargetMode("External");
		rels.addRelationship(rel);
		link.setId(id);
		links.add(link);
	}

	public void createComment(WorksheetBuilder worksheet, Cell cell, String author, String comment, CommentPosition position) throws Docx4JException {
		int sheetIndex = sheets.indexOf(worksheet.sheet);
		CTComments comments = null;
		Xml drawings = null;
		if (this.comments.size() <= sheetIndex || this.comments.get(sheetIndex) == null) {
			CommentsPart commentsPart = new CommentsPart(new PartName("/xl/comments" + (sheetIndex+1) + ".xml"));
			while (this.comments.size() <= sheetIndex) {
				this.comments.add(null);
			}
			this.comments.set(sheetIndex, commentsPart);
			//String commentsId = worksheet.sheet.addTargetPart(commentsPart).getId();
			
			comments = new CTComments();
			commentsPart.setContents(comments);
			comments.setAuthors(new CTAuthors());
			comments.setCommentList(new CTCommentList());
			
			//Also create a 'drawing' for the comments
			VMLPart vml = new VMLPart();
			drawings = new Xml();
			vml.setContents(drawings);
			//o:shapelayout
			CTShapeLayout layout = new CTShapeLayout();
			layout.setExt(STExt.EDIT);
			CTIdMap map = new CTIdMap();
			map.setExt(STExt.EDIT);
			map.setData("1");
			layout.setIdmap(map);
			addJAXBElement(layout, drawings.getAny(), "urn:schemas-microsoft-com:office:office", "shapelayout", "o", CTShapeLayout.class);
					
			//v:shapetype
			CTShapetype type = new CTShapetype();
			type.setPath("m,l,21600r21600,l21600,xe");
			type.setSpt(202f);
			type.setCoordsize("21600,21600");
			type.setVmlId("_x0000_t202");
			CTStroke stroke = new CTStroke();
			stroke.setJoinstyle(STStrokeJoinStyle.MITER);
			addJAXBElement(stroke, type.getEGShapeElements(), CTStroke.class, "urn:schemas-microsoft-com:vml", "stroke", "v");
			CTPath path = new CTPath();
			path.setConnecttype(STConnectType.RECT);
			path.setGradientshapeok(STTrueFalse.T);
			addJAXBElement(path, type.getEGShapeElements(), CTPath.class, "urn:schemas-microsoft-com:vml", "path", "v");
			
			addJAXBElement(type, drawings.getAny(), "urn:schemas-microsoft-com:vml", "shapetype", "v", CTShapetype.class);
			
			while (this.commentsDrawings.size() <= sheetIndex) {
				this.commentsDrawings.add(null);
			}
			this.commentsDrawings.set(sheetIndex, vml);
			String drawingId = worksheet.sheet.addTargetPart(vml).getId();
			CTLegacyDrawing legacy = new CTLegacyDrawing();
			legacy.setId(drawingId);
			worksheet.sheet.getContents().setLegacyDrawing(legacy);
		} else {
			comments = this.comments.get(sheetIndex).getContents();
			drawings = this.commentsDrawings.get(sheetIndex).getContents();
		}

		//Make the actual comment
		List<String> authors = comments.getAuthors().getAuthor();
		int authorId = authors.size();
		for (int i=0; i<authors.size(); i++) {
			if (authors.get(i).equals(author)) {
				authorId = i;
				break;
			}
		}
		if (authorId == authors.size()) {
			authors.add(author);
		}
		
		CTComment newComment = new CTComment();
		newComment.setAuthorId(authorId);
		newComment.setRef(cell.getR());
		newComment.setText(createCommentText(author, comment));
		
		comments.getCommentList().getComment().add(newComment);
		
		//And create the drawing shape for it...
		CTShape shape = new CTShape();
		shape.setVmlId("_x0000_s" + (int)Math.floor(Math.random() * 10000));
		shape.setFillcolor("#ffffe1");
		shape.setStyle("position:absolute;margin-left:59.25pt;margin-top:59.25pt;width:108pt;height:59.25pt;z-index:1");
		shape.setType("#_x0000_t202");
		shape.setInsetmode(STInsetMode.AUTO);
		addJAXBElement(shape, drawings.getAny(), "urn:schemas-microsoft-com:vml", "shape", "v", CTShape.class);

		org.docx4j.vml.CTFill fill = new org.docx4j.vml.CTFill();
		fill.setColor2("#ffffe1");
		addJAXBElement(fill, shape.getEGShapeElements(), org.docx4j.vml.CTFill.class, "urn:schemas-microsoft-com:vml", "fill", "v");
		
		CTShadow shadow = new CTShadow();
		shadow.setObscured(STTrueFalse.T);
		shadow.setColor("black");
		addJAXBElement(shadow, shape.getEGShapeElements(), CTShadow.class, "urn:schemas-microsoft-com:vml", "shadow", "v");
		
		CTPath path = new CTPath();
		path.setConnecttype(STConnectType.NONE);
		addJAXBElement(path, shape.getEGShapeElements(), CTPath.class, "urn:schemas-microsoft-com:vml", "path", "v");
		
		try {
			JAXBContext jc = JAXBContext.newInstance(
					org.docx4j.w14.ObjectFactory.class,
					org.docx4j.w15.ObjectFactory.class,
					org.docx4j.wml.ObjectFactory.class,
					org.docx4j.vml.spreadsheetDrawing.ObjectFactory.class,
					org.docx4j.vml.ObjectFactory.class,
					org.docx4j.vml.officedrawing.ObjectFactory.class,
					org.docx4j.vml.wordprocessingDrawing.ObjectFactory.class,
					org.docx4j.vml.presentationDrawing.ObjectFactory.class);

			JAXBElement<CTTextbox> box = jc.createUnmarshaller().unmarshal(new StreamSource(new StringReader("<textbox style=\"mso-direction-alt:auto\"><div style=\"text-align: left\" /></textbox>")), CTTextbox.class);
			addJAXBElement(box.getValue(), shape.getEGShapeElements(), CTTextbox.class, "urn:schemas-microsoft-com:vml", "path", "v");
			
			XLSXRange cellRange = XLSXRange.fromCell(cell);
			String clientData = "<ClientData ObjectType=\"Note\">" +
					"<MoveWithCells />" +
					"<SizeWithCells />" +
					"<Anchor>" + position.toString() + "</Anchor>" +
					"<AutoFill>False</AutoFill>" +
					"<Row>" + cellRange.startCellNumericRow() + "</Row>" +
					"<Column>" + cellRange.startCellNumericColumn() + "</Column>" +
					"<Visible />" +
				"</ClientData>";
				
				
			JAXBElement<CTClientData> data = jc.createUnmarshaller().unmarshal(new StreamSource(new StringReader(clientData)), CTClientData.class);
			addJAXBElement(data.getValue(), shape.getEGShapeElements(), CTClientData.class, "urn:schemas-microsoft-com:office:excel", "ClientData", "x");
		} catch (Exception e) {
			e.printStackTrace();
		}

		
	}

	private <T> void addJAXBElement(T layout, List<Object> any, String ns, String local, String prefix, Class<T> clazz) {
		any.add(new JAXBElement<T>(new QName(ns, local, prefix), clazz, layout));
	}

	private <T> void addJAXBElement(T layout, List<JAXBElement<?>> any, Class<T> clazz, String ns, String local, String prefix) {
		any.add(new JAXBElement<T>(new QName(ns, local, prefix), clazz, layout));
	}

	private CTRst createCommentText(String author, String comment) {
		CTRst text = new CTRst();
		text.getR().add(createCommentChunk(author + ": ", "<rPr><b/><sz val=\"9\"/><color indexed=\"81\" /><rFont val=\"Tahoma\"/><family val=\"2\"/></rPr>"));
		text.getR().add(createCommentChunk(" " + comment, "<rPr><sz val=\"9\"/><color indexed=\"81\" /><rFont val=\"Tahoma\"/><family val=\"2\"/></rPr>"));
		return text;
	}

	private CTRElt createCommentChunk(String value, String rpr) {
		CTRElt r = null;
		try {
			JAXBContext jc = JAXBContext.newInstance(CTRPrElt.class);
			JAXBElement<CTRElt> element = jc.createUnmarshaller().unmarshal(new StreamSource(new StringReader("<r>" + rpr + "</r>")), CTRElt.class); 
			r = element.getValue();
		} catch (JAXBException e) {
			e.printStackTrace();
			r = new CTRElt();
		}
		CTXstringWhitespace text = new CTXstringWhitespace();
		text.setValue(value);
		text.setSpace("preserve");
		r.setT(text);
		return r;
	}

	public void installThickBottomStyle(int index) {
		this.thickBottomStyles.add(new Long(index));
	}

	public boolean isThickBottomStyle(Long style) {
		return thickBottomStyles.contains(style);
	}

	public void createPageSetup(WorksheetBuilder worksheet, InputStream data, STOrientation orientation, int scale, int paperSize) throws Docx4JException {
		int sheetIndex = sheets.indexOf(worksheet.sheet);
		PrinterSettings settings = new PrinterSettings(new PartName("/xl/printerSettings/printerSettings" + sheetIndex + ".bin"));
		settings.setBinaryData(data);
		
		String id = worksheet.sheet.addTargetPart(settings).getId();
		
		CTPageSetup setup = new CTPageSetup();
		setup.setOrientation(orientation);
		setup.setScale(new Long(scale));
		setup.setPaperSize(new Long(paperSize));
		setup.setId(id);
		
		worksheet.sheet.getContents().setPageSetup(setup);
	}

	/**
	 * Provides support for unlocked cells in a locked worksheet.
	 * 
	 * The style with the given id is cloned but with the cell protection set to unlocked.
	 */
	public Long getUnlockedStyle(long styleId) {
		Long id = unlockedStyles.get(styleId);
		if (id != null) 
			return id;
		
		//Find the original style
		CTXf originalStyle = stylesheet.getCellXfs().getXf().get((int)styleId);
		CTXf xf = new CTXf();
		xf.setNumFmtId(originalStyle.getNumFmtId());
		xf.setFontId(originalStyle.getFontId());
		xf.setFillId(originalStyle.getFillId());
		xf.setBorderId(originalStyle.getBorderId());
		xf.setApplyNumberFormat(originalStyle.isApplyNumberFormat());
		xf.setApplyFont(originalStyle.isApplyFont());
		xf.setApplyFill(originalStyle.isApplyFill());
		xf.setApplyBorder(originalStyle.isApplyBorder());
		//Always ref the default style
		xf.setXfId(0L);
		
		xf.setAlignment(originalStyle.getAlignment());
		xf.setApplyProtection(true);
		
		CTCellProtection protection = new CTCellProtection();
		protection.setLocked(false);
		xf.setProtection(protection);
		
		long index = stylesheet.getCellXfs().getXf().size();
		stylesheet.getCellXfs().getXf().add(xf);
		
		unlockedStyles.put(styleId, index);
		return index;
	}

	int createMultiStyleText() {
		int index = strings.getSi().size();
		CTRst si = new CTRst();
		strings.getSi().add(si);
		//TODO Cache it???
//		stringCache.put(value, index);
		return index;
	}

	CTRst getMultiStyleString(int index) {
		return strings.getSi().get(index);
	}

	JAXBElement<?> createFontModification(STVerticalAlignRun type) {
		CTVerticalAlignFontProperty vertAlign = new CTVerticalAlignFontProperty();
		vertAlign.setVal(type);
		return factory.createCTRPrEltVertAlign(vertAlign);
	}
	
	SpreadsheetMLPackage getSpreadsheetMLPackage() {
		return pkg;
	}
}
