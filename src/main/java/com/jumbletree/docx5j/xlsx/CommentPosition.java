package com.jumbletree.docx5j.xlsx;

public class CommentPosition {

	private int topRow;
	private int topOffset;
	private int leftColumn;
	private int leftOffset;

	private int bottomRow;
	private int bottomOffset;
	private int rightColumn;
	private int rightOffset;
	
	
	
	public CommentPosition(int leftColumn, int leftOffset, int topRow, int topOffset, int rightColumn, int rightOffset, int bottomRow, int bottomOffset) {
		super();
		this.topRow = topRow;
		this.topOffset = topOffset;
		this.leftColumn = leftColumn;
		this.leftOffset = leftOffset;
		this.bottomRow = bottomRow;
		this.bottomOffset = bottomOffset;
		this.rightColumn = rightColumn;
		this.rightOffset = rightOffset;
	}
	
	public int getTopRow() {
		return topRow;
	}
	public void setTopRow(int topRow) {
		this.topRow = topRow;
	}
	public int getTopOffset() {
		return topOffset;
	}
	public void setTopOffset(int topOffset) {
		this.topOffset = topOffset;
	}
	public int getLeftColumn() {
		return leftColumn;
	}
	public void setLeftColumn(int leftColumn) {
		this.leftColumn = leftColumn;
	}
	public int getLeftOffset() {
		return leftOffset;
	}
	public void setLeftOffset(int leftOffset) {
		this.leftOffset = leftOffset;
	}
	public int getBottomRow() {
		return bottomRow;
	}
	public void setBottomRow(int bottomRow) {
		this.bottomRow = bottomRow;
	}
	public int getBottomOffset() {
		return bottomOffset;
	}
	public void setBottomOffset(int bottomOffset) {
		this.bottomOffset = bottomOffset;
	}
	public int getRightColumn() {
		return rightColumn;
	}
	public void setRightColumn(int rightColumn) {
		this.rightColumn = rightColumn;
	}
	public int getRightOffset() {
		return rightOffset;
	}
	public void setRightOffset(int rightOffset) {
		this.rightOffset = rightOffset;
	}

	@Override
	public String toString() {
		return leftColumn + ", " + leftOffset + ", " + topRow + ", " + topOffset + ", " + rightColumn + ", " + rightOffset + ", " + bottomRow + ", " + bottomOffset;
	}

}
