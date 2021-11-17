package com.jumbletree.docx5j.xlsx.builders;

import java.awt.Color;
import java.util.ArrayList;
import java.util.List;

import org.docx4j.sharedtypes.STVerticalAlignRun;

import com.jumbletree.docx5j.xlsx.builders.text.TextModifier;

public class TextModifierBuilder {

	public static TextModifierBuilder newBuilder() {
		return new TextModifierBuilder();
	}
	
	List<TextModifier> mods = new ArrayList<>();
	
	public TextModifierBuilder withSuperscript() {
		mods.add((o,l)->l.add(o.createFontModification(STVerticalAlignRun.SUPERSCRIPT)));
		return this;
	}

	public TextModifierBuilder withSubscript() {
		mods.add((o,l)->l.add(o.createFontModification(STVerticalAlignRun.SUBSCRIPT)));
		return this;
	}
	
	public TextModifierBuilder withBaseline() {
		mods.add((o,l)->l.add(o.createFontModification(STVerticalAlignRun.BASELINE)));
		return this;
	}

	public TextModifierBuilder withColor(Color color) {
		mods.add((o,l)->l.add(o.createFontModification(color)));
		return this;
	}
	
	public TextModifier build() {
		return (o,l)->{
			mods.forEach(mod->mod.apply(o, l));
		};
	}
}
