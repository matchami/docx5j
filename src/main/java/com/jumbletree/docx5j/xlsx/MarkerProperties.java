package com.jumbletree.docx5j.xlsx;

import java.awt.Color;

import org.docx4j.dml.chart.STMarkerStyle;

/**
 * Markers currently default to no-line
 * @author matchami
 *
 */
public class MarkerProperties {

	private STMarkerStyle style;
	private Color color;

	public MarkerProperties(STMarkerStyle style, Color color) {
		this.style = style;
		this.color = color;
	}

	public STMarkerStyle getStyle() {
		return style;
	}

	public Color getColor() {
		return color;
	}
	
}
