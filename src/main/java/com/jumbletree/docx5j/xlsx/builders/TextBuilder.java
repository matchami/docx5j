package com.jumbletree.docx5j.xlsx.builders;

import java.util.List;
import java.util.function.BiConsumer;

import javax.xml.bind.JAXBElement;

import org.docx4j.sharedtypes.STVerticalAlignRun;
import org.xlsx4j.sml.CTRElt;
import org.xlsx4j.sml.CTRPrElt;
import org.xlsx4j.sml.CTRst;
import org.xlsx4j.sml.CTXstringWhitespace;

import com.jumbletree.docx5j.xlsx.builders.text.TextModifier;

public class TextBuilder {

	public static enum Modifier implements TextModifier {
		NO_MODIFICATION((o,l)->{}),
		SUPERSCRIPT((o,l)->l.add(o.createFontModification(STVerticalAlignRun.SUPERSCRIPT))),
		SUBSCRIPT((o,l)->l.add(o.createFontModification(STVerticalAlignRun.SUBSCRIPT))),
		BASELINE((o,l)->l.add(o.createFontModification(STVerticalAlignRun.BASELINE)));
		
		private BiConsumer<WorkbookBuilder, List<JAXBElement<?>>> mod;

		Modifier(BiConsumer<WorkbookBuilder, List<JAXBElement<?>>> mod) {
			this.mod = mod;
		}
		
		public void apply(WorkbookBuilder origin, List<JAXBElement<?>> list) {
			mod.accept(origin, list);
		}
	}

	private CellBuilder parent;
	private WorkbookBuilder origin;
	private int index;
	private int baseStyle;

	TextBuilder(CellBuilder parent, WorkbookBuilder origin, long baseStyle, int index) {
		this.parent = parent;
		this.origin = origin;
		this.index = index;
		this.baseStyle = (int)baseStyle;
	}

	public TextBuilder add(String text) {
		return add(text, Modifier.NO_MODIFICATION);
	}
	
	public TextBuilder add(String text, TextModifier mod) {
		CTRst si = origin.getMultiStyleString(index);
		List<CTRElt> rs = si.getR();
		
		CTRElt r = new CTRElt();
		CTXstringWhitespace t = new CTXstringWhitespace();
		t.setValue(text);
		if (text.startsWith(" ") || text.endsWith(" ")) {
			t.setSpace("preserve");
		}
		r.setT(t);

		if (rs.size() == 0 && mod == Modifier.NO_MODIFICATION) {
			//No formatting necessary, just add it
		} else {
			CTRPrElt rpr = new CTRPrElt();
			List<JAXBElement<?>> font = rpr.getRFontOrCharsetOrFamily();
			
			//Find current props from style
			List<JAXBElement<?>> baseFont = origin.getStyleFont(baseStyle).getNameOrCharsetOrFamily();
			
			font.addAll(baseFont);
			
			//And apply the mod
			mod.apply(origin, font);
			r.setRPr(rpr);
		}
		
		rs.add(r);
		return this;
	}
	
	public CellBuilder cell() {
		return parent;
	}
}
