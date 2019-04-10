package com.jumbletree.docx5j.xlsx;

import org.xlsx4j.sml.STSheetViewType;
import org.xlsx4j.sml.SheetView;

public class View {

	private int workbookViewId;
	private Long zoomScaleSheetLayoutView;
	private Long zoomScaleNormal;
	private Long zoomScale;
	private STSheetViewType view;
	private Boolean tabSelected;
	public View(int workbookViewId, int zoomScaleSheetLayoutView, int zoomScaleNormal, int zoomScale, STSheetViewType view, boolean tabSelected) {
		this.workbookViewId = workbookViewId;
		this.zoomScaleSheetLayoutView = Long.valueOf(zoomScaleSheetLayoutView);
		this.zoomScaleNormal = Long.valueOf(zoomScaleNormal);
		this.zoomScale = Long.valueOf(zoomScale);
		this.view = view;
		this.tabSelected = tabSelected;
	}
	
	public View(int workbookViewId) {
		this.workbookViewId = workbookViewId;
	}

	public int getWorkbookViewId() {
		return workbookViewId;
	}
	public Long getZoomScaleSheetLayoutView() {
		return zoomScaleSheetLayoutView;
	}
	public Long getZoomScaleNormal() {
		return zoomScaleNormal;
	}
	public Long getZoomScale() {
		return zoomScale;
	}
	public STSheetViewType getView() {
		return view;
	}
	public boolean getTabSelected() {
		return tabSelected;
	}

	public SheetView createSheetView() {
		SheetView view = new SheetView();
		view.setWorkbookViewId(workbookViewId);
		view.setZoomScaleSheetLayoutView(zoomScaleSheetLayoutView);
		view.setZoomScaleNormal(zoomScaleNormal);
		view.setZoomScale(zoomScale);
		view.setView(this.view);
		view.setTabSelected(tabSelected);
		return view;
	}
	
	
}
