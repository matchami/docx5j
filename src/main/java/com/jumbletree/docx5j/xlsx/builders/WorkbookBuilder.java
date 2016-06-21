package com.jumbletree.docx5j.xlsx.builders;

import java.awt.Color;
import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import javax.xml.bind.JAXBElement;
import javax.xml.bind.JAXBException;

import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.exceptions.InvalidFormatException;
import org.docx4j.openpackaging.io3.Save;
import org.docx4j.openpackaging.packages.SpreadsheetMLPackage;
import org.docx4j.openpackaging.parts.PartName;
import org.docx4j.openpackaging.parts.SpreadsheetML.SharedStrings;
import org.docx4j.openpackaging.parts.SpreadsheetML.Styles;
import org.docx4j.openpackaging.parts.SpreadsheetML.WorksheetPart;
import org.xlsx4j.jaxb.Context;
import org.xlsx4j.sml.CTBooleanProperty;
import org.xlsx4j.sml.CTBorder;
import org.xlsx4j.sml.CTBorderPr;
import org.xlsx4j.sml.CTBorders;
import org.xlsx4j.sml.CTCellStyle;
import org.xlsx4j.sml.CTCellStyleXfs;
import org.xlsx4j.sml.CTCellStyles;
import org.xlsx4j.sml.CTCellXfs;
import org.xlsx4j.sml.CTColor;
import org.xlsx4j.sml.CTFill;
import org.xlsx4j.sml.CTFills;
import org.xlsx4j.sml.CTFont;
import org.xlsx4j.sml.CTFontFamily;
import org.xlsx4j.sml.CTFontName;
import org.xlsx4j.sml.CTFontScheme;
import org.xlsx4j.sml.CTFontSize;
import org.xlsx4j.sml.CTFonts;
import org.xlsx4j.sml.CTPatternFill;
import org.xlsx4j.sml.CTRst;
import org.xlsx4j.sml.CTSst;
import org.xlsx4j.sml.CTStylesheet;
import org.xlsx4j.sml.CTXf;
import org.xlsx4j.sml.CTXstringWhitespace;
import org.xlsx4j.sml.Cell;
import org.xlsx4j.sml.ObjectFactory;
import org.xlsx4j.sml.Row;
import org.xlsx4j.sml.STCellType;
import org.xlsx4j.sml.STFontScheme;
import org.xlsx4j.sml.STPatternType;
import org.xlsx4j.sml.Sheet;
import org.xlsx4j.sml.SheetData;

import com.jumbletree.docx5j.xlsx.LineChart;
import com.jumbletree.docx5j.xlsx.ScatterChart;
import com.jumbletree.docx5j.xlsx.XLSXRange;

public class WorkbookBuilder implements BuilderMethods {

	private SpreadsheetMLPackage pkg;
	private ArrayList<WorksheetPart> sheets;
	private CTStylesheet stylesheet;
	private ObjectFactory factory;
	private CTSst strings;
	private HashMap<String, Integer> stringCache;
	private HashMap<String, Long> styles = new HashMap<>();

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
		createFont("Calibri", 11, Color.black, false, false);
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
		border.setBottom(new CTBorderPr());
		border.setDiagonal(new CTBorderPr());
		stylesheet.setBorders(new CTBorders());
		stylesheet.getBorders().getBorder().add(border);
	}

	protected void createDefaultFill() {
		CTFills fills = new CTFills();
		stylesheet.setFills(fills);
		CTFill fill = new CTFill();
		CTPatternFill pattern = new CTPatternFill();
		pattern.setPatternType(STPatternType.NONE);
		fill.setPatternFill(pattern);
		fills.getFill().add(fill);
	}

	int createStyle(Long formatId, Long fontId, Long fillId, Long borderId) {
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
		
		int index = stylesheet.getCellXfs().getXf().size();
		stylesheet.getCellXfs().getXf().add(xf);
		
		return index;
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
	
	public int createBorder() {
		//TODO
		return 0;
	}
	
	int createFont(String fontName, int size, Color color, boolean bold, boolean italic) {
		CTFont font = new CTFont();
		setFontSize(size, font);
		setFontName(fontName, font);
		setFontColor(color, font);
		if (bold) {
			CTBooleanProperty bool = new CTBooleanProperty();
			bool.setVal(true);
			font.getNameOrCharsetOrFamily().add(factory.createCTFontB(bool));
		}
		if (italic) {
			CTBooleanProperty bool = new CTBooleanProperty();
			bool.setVal(true);
			font.getNameOrCharsetOrFamily().add(factory.createCTFontI(bool));
		}
		CTFontFamily family = new CTFontFamily();
		family.setVal(2);
		font.getNameOrCharsetOrFamily().add(factory.createCTFontFamily(family));
		CTFontScheme scheme = new CTFontScheme();
		scheme.setVal(STFontScheme.MINOR);
		font.getNameOrCharsetOrFamily().add(factory.createCTFontScheme(scheme));
		if (stylesheet.getFonts() == null) {
			stylesheet.setFonts(new CTFonts());
		}
		int index = stylesheet.getFonts().getFont().size();
		stylesheet.getFonts().getFont().add(font);
		
		return index;
	}
	
	private void setFontSize(long size, CTFont font){
		CTFontSize fontSize = new CTFontSize();
		fontSize.setVal(size);
		fontSize.setParent(font);
		font.getNameOrCharsetOrFamily().add(factory.createCTFontSz(fontSize));
	}

	private void setFontColor(Color color, CTFont font){
		CTColor fontCol = new CTColor();
		fontCol.setRgb(getColorBytes(color));
		fontCol.setTheme( new Long(1) );
		fontCol.setTint( new Double(0.0) );
		fontCol.setParent(font);
		JAXBElement<CTColor> element1 = factory.createCTFontColor(fontCol);
		font.getNameOrCharsetOrFamily().add(element1);
	}

	private void setFontName(String name, CTFont font){
		CTFontName fontName = new CTFontName();
		fontName.setVal(name);
		fontName.setParent(font);
		font.getNameOrCharsetOrFamily().add(factory.createCTFontName(fontName));
	}

	public void setValue(int sheet, int col, int row, double value) throws Docx4JException {
		setValue(sheet, col, row, value, null);
	}
	
	public void setValue(int sheet, int col, int row, double value, Long style) throws Docx4JException {
		Cell theCell = getCell(sheet, col, row);

		theCell.setV(String.valueOf(value));
		if (style != null)
			theCell.setS(style);
	}

	public void setValue(int sheet, int col, int row, String value) throws Docx4JException {
		setValue(sheet, col, row, value, null);
	}

	public void setValue(int sheet, int col, int row, String value, String styleName) throws Docx4JException {
		Cell theCell = getCell(sheet, col, row);

		theCell.setT(STCellType.S);

		int index = getStringCache(value);
		theCell.setV(String.valueOf(index));
		if (styleName != null)
			theCell.setS(styles.get(styleName));
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
	 * 
	 * @param series a list of name, range pairs that will make up the chart eg createChart("First series", "Sheet1!$B$1:$B$20");
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
}
