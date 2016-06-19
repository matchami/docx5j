package com.jumbletree.docx5j.xlsx;

import java.awt.Color;

import org.docx4j.dml.CTSRgbColor;
import org.docx4j.dml.chart.CTBoolean;
import org.docx4j.dml.chart.CTDouble;
import org.docx4j.dml.chart.CTUnsignedInt;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.SpreadsheetMLPackage;
import org.docx4j.openpackaging.parts.PartName;
import org.docx4j.openpackaging.parts.Parts;

public interface FactoryMethods {


	public default <T> T createVal(Class<T> clazz, Object value) throws Docx4JException {
		try { 
			T t = clazz.newInstance();
			clazz.getMethod("setVal", value.getClass()).invoke(t, value);
			return t;
		} catch (Exception e) {
			throw new Docx4JException("Incompatible value: " + clazz.getName() + "/" + value.getClass().getName(), e);
		}
	}

	public default CTUnsignedInt createUnsignedInt(int value) {
		CTUnsignedInt unsign1 = new CTUnsignedInt();
	    unsign1.setVal(value);
	    return unsign1;
	}

	public default CTDouble createDouble(double d) {
		CTDouble dbl = new CTDouble();
		dbl.setVal(d);
		return dbl;
	}

	public default CTBoolean createBoolean(boolean flag) {
		CTBoolean bool = new CTBoolean();
		bool.setVal(flag);
		return bool;
	}
	
	public default int getNextPartNumber(String prefix, SpreadsheetMLPackage pkg) {
		int chartNumber = 1;
		Parts parts = pkg.getParts();
		for (PartName name : parts.getParts().keySet()) {
			if (name.getName().startsWith(prefix)) {
				int num = Integer.parseInt(name.getName().substring(prefix.length(), name.getName().indexOf(".")));
				chartNumber = Math.max(chartNumber, num+1);
			}
		}
		return chartNumber;
	}
	
	public default CTSRgbColor createColor(Color color) {
		CTSRgbColor rgb = new CTSRgbColor();
		rgb.setVal(getColorString(color));
		return rgb;
	}

	public default byte[] getColorBytes(Color color) {
		return new byte[] {(byte)color.getRed(), (byte)color.getGreen(), (byte)color.getBlue()};
	}
	
	public default String getColorString(Color color) {
		return getPaddedHex(color.getRed()) +
				getPaddedHex(color.getGreen()) +
				getPaddedHex(color.getBlue());
	}

	public default String getPaddedHex(int val) {
		String hex = Integer.toHexString(val);
		if (hex.length() == 1) {
			hex = "0" + hex;
		}
		return hex;
	}
}
