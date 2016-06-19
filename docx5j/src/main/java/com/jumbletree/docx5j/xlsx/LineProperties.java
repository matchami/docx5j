package com.jumbletree.docx5j.xlsx;

import java.awt.Color;

public class LineProperties {

	int width = 28800;
	Color color;
	
	public LineProperties(Color color) {
		this.color = color;
	}
	
	public LineProperties(Color color, int lineWidth) {
		this.width = lineWidth;
		this.color = color;
	}

	public int getWidth() {
		return width;
	}

	public Color getColor() {
		return color;
	}
}
