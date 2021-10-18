package com.jumbletree.docx5j.xlsx.builders;

import java.util.HashMap;
import java.util.Map;

import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.xlsx4j.sml.STSheetViewType;
import org.xlsx4j.sml.SheetView;

public class SheetViewBuilder {

	public static String TAB_SELECTED = "TabSelected";
	public static String SHOW_GRID_LINES = "ShowGridLines";
	
	private int workbookViewId = 0;
	private Long zoomScaleSheetLayoutView = 100l;
	private Long zoomScaleNormal = 100l;
	private Long zoomScale = 100l;
	private STSheetViewType view;
	private HashMap<String, Boolean> properties = new HashMap<>();
	
	public SheetViewBuilder(STSheetViewType view) {
		this.view = view;
		
		properties.put(TAB_SELECTED, Boolean.TRUE);
	}

	public static SheetViewBuilder newBuilder(STSheetViewType view) {
		return new SheetViewBuilder(view);
	}
	
	public SheetViewBuilder withZoom(long zoom) {
		this.zoomScale = zoom;
		return this;
	}
	
	public SheetViewBuilder withProperty(String key, boolean value) {
		this.properties.put(key, value);
		return this;
	}

	public void addTo(WorksheetBuilder sheet) throws Docx4JException {
		sheet.addView(build()); 
	}

	private SheetView build() {
		SheetView view = new SheetView();
		view.setWorkbookViewId(workbookViewId);
		view.setZoomScaleSheetLayoutView(zoomScaleSheetLayoutView);
		view.setZoomScaleNormal(zoomScaleNormal);
		view.setZoomScale(zoomScale);
		view.setView(this.view);
		for (Map.Entry<String, Boolean> prop : properties.entrySet()) {
			String methodName = "set" + prop.getKey();
			try {
				SheetView.class.getMethod(methodName, Boolean.class)
					.invoke(view, prop.getValue());
			} catch (Exception e) {
				//TODO Log properly
				e.printStackTrace();
			}
		}
		return view;
	}

}
