package com.jumbletree.docx5j.xlsx;

import java.util.ArrayList;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.xlsx4j.sml.Cell;

public class XLSXRange {
	static Pattern pattern = Pattern.compile("([A-Z]+)([0-9]+)");

	String sheet;
	String startCell;
	String endCell;
	public XLSXRange(String startCell, String endCell) {
		this(null, startCell, endCell);
	}
	public XLSXRange(String sheet, String startCell, String endCell) {
		this.sheet = sheet;
		this.startCell = startCell.toUpperCase();
		this.endCell = endCell.toUpperCase();
	}
	public String getSheet() {
		return sheet;
	}
	public String getStartCell() {
		return startCell;
	}
	public String getEndCell() {
		return endCell;
	}
	public String singleCellAbsoluteReference() {
		return sheet + "!" + absolute(startCell);
	}
	public String rangeAbsoluteReference() {
		return sheet + "!" + absolute(startCell) + ":" + absolute(endCell);
	}
	public String singleCellSheetlessReference() {
		return startCell;
	}
	public String rangeSheetlessReference() {
		return startCell + ":" + endCell;
	}
	private String absolute(String cellRef) {
		Matcher matcher = pattern.matcher(cellRef);
		if (!matcher.find()) {
			throw new IllegalArgumentException("Invalid cell reference given");
		}
		return "$" + matcher.group(1) + "$" + matcher.group(2);
	}
//	
//	public static void main(String[] args) {
//		System.out.println(new XLSXRange("Sheet1", "B4", "D4").singleCellAbsoluteReference());
//	}
	public int startCellNumericColumn() {
		return numericColumn(startCell);
	}
	public int startCellNumericRow() {
		return numericRow(startCell);
	}
	public int endCellNumericColumn() {
		return numericColumn(endCell);
	}
	public int endCellNumericRow() {
		return numericRow(endCell);
	}
	/**
	 * Row refs appear to be zero based
	 * @param cellRef
	 * @return
	 */
	private int numericRow(String cellRef) {
		Matcher matcher = pattern.matcher(cellRef);
		if (!matcher.find()) {
			throw new IllegalArgumentException("Invalid cell reference given");
		}
		return Integer.parseInt(matcher.group(2)) - 1;
	}
	
	/**
	 * Col refs appear to be zero based
	 * @param cellRef
	 * @return
	 */
	private int numericColumn(String cellRef) {
		Matcher matcher = pattern.matcher(cellRef);
		if (!matcher.find()) {
			throw new IllegalArgumentException("Invalid cell reference given");
		}
		String col = matcher.group(1);
		int power = col.length() - 1;
		int column = 0;
		while (power >= 0) {
			//A=1, B=2 etc
			column += " ABCDEFGHIJKLMNOPQRSTUVWXYZ".indexOf(col.charAt(0)) * Math.pow(26, power);
			col = col.substring(1);
			--power;
		}
		//And now make it zero based
		return column - 1;
	}
	
	public static String asCell(int col, int row) {
		String str = getStr(col);
		return str + (row+1);
	}
	
	private static String getStr(int col) {
		if (col >= 26) {
			int colBit = col % 26;
			System.out.println("Bit: " + colBit);
			col = col - colBit;
			col = col / 26;
			col -= 1;
			System.out.println("Col: " + col);
			return getStr(col) + getStr(colBit);
		} else {
			return String.valueOf("ABCDEFGHIJKLMNOPQRSTUVWXYZ".charAt(col));
		}
	}
	
	public int getLinearSize() {
		if (isLinearDimensionHorizontal()) {
			return numericColumn(endCell) - numericColumn(startCell) + 1;
		} else {
			return numericRow(endCell) - numericRow(startCell) + 1;
		}
	}
	
	private boolean isLinearDimensionHorizontal() {
		if (numericColumn(startCell) == numericColumn(endCell)) {
			return false;
		}
		if (numericRow(startCell) == numericRow(endCell)) {
			return true;
		}
		throw new IllegalArgumentException("Range is linear in neither horizontal or vertical direction");
		
	}
	
	public List<XLSXRange> getLinearCells() {
		List<XLSXRange> list = new ArrayList<XLSXRange>();
		if (isLinearDimensionHorizontal()) {
			int row = numericRow(startCell);
			for (int i=numericColumn(startCell); i<=numericColumn(endCell); i++) {
				list.add(fromNumericReference(sheet, i, row));
			}
		} else {
			int col = numericColumn(startCell);
			for (int i=numericRow(startCell); i<=numericRow(endCell); i++) {
				list.add(fromNumericReference(sheet, col, i));
			}
		}
		return list;
	}
	
	public static XLSXRange fromNumericReference(String sheet, int col, int row) {
		return singleCell(sheet, getStr(col) + (row + 1));
	}

	public static XLSXRange fromNumericReference(int col, int row) {
		return singleCell(null, getStr(col) + (row + 1));
	}

	private static XLSXRange singleCell(String sheet, String cell) {
		return new XLSXRange(sheet, cell, cell);
	}
	public static XLSXRange fromReference(String value) {
		String[] bits = value.split("[\\!\\:]");
		if (bits.length == 1) {
			return singleCell(null, bits[0].replaceAll("\\$", ""));
		} else if (bits.length == 2) {
			return singleCell(bits[0], bits[1].replaceAll("\\$", ""));
		}
		return new XLSXRange(bits[0], bits[1].replaceAll("\\$", ""), bits[2].replaceAll("\\$", ""));
	}
	public String absoluteReference() {
		return isSingleCell() ? singleCellAbsoluteReference() : rangeAbsoluteReference();
	}
	private boolean isSingleCell() {
		return endCell.equals(startCell);
	}
	public static XLSXRange fromCell(Cell cell) {
		return fromReference(cell.getR());
	}
}
