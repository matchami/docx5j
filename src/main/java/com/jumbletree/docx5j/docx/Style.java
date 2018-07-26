package com.jumbletree.docx5j.docx;

import java.math.BigInteger;

import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.openpackaging.parts.WordprocessingML.StyleDefinitionsPart;
import org.docx4j.wml.CTBorder;
import org.docx4j.wml.Color;
import org.docx4j.wml.HpsMeasure;
import org.docx4j.wml.Jc;
import org.docx4j.wml.JcEnumeration;
import org.docx4j.wml.ObjectFactory;
import org.docx4j.wml.PPr;
import org.docx4j.wml.PPrBase.Spacing;
import org.docx4j.wml.STBorder;
import org.docx4j.wml.PPrBase.PBdr;
import org.docx4j.wml.RPr;
import org.docx4j.wml.Style.BasedOn;

public class Style {

	private String css;
	private ObjectFactory factory;

	public Style(String css, ObjectFactory factory) {
		this.css = css;
		this.factory = factory;
	}

	public void addTo(MainDocumentPart main, String styleName, String id, String basedOn) {
		
		StyleDefinitionsPart styles = main.getStyleDefinitionsPart();

		org.docx4j.wml.Style style = factory.createStyle();
		style.setType("paragraph");
		style.setStyleId(id);
		
		org.docx4j.wml.Style.Name name = factory.createStyleName();
		name.setVal(styleName);
	    style.setName(name);
	    
	    BasedOn based = factory.createStyleBasedOn();
	    based.setVal("Normal");      
	    style.setBasedOn(based);

	    PPr ppr = null;
	    PBdr border = null;
	    RPr rpr = null;
	    Spacing spacing = null;
	    
	    //Now do the actual style
	    for (String definition : css.split("[{};]+")) {
	    	definition = definition.trim();
	    	if (definition.length() == 0) 
	    		continue;
	    	
	    	int index = definition.indexOf(":");
	    	if (index == -1)
	    		continue;
	    	
	    	String command = definition.substring(0, index);
	    	String value = definition.substring(index+1).trim();
	    	
	    	if (command.equals("color")) {
	    		if (rpr == null) {
	    			rpr = factory.createRPr();
	    			style.setRPr(rpr);
	    		}	    		
	    	    Color color = factory.createColor();
	    	    if (!value.startsWith("#"))
	    	    	throw new IllegalArgumentException("Only RGB colors supported for now");
	    	    color.setVal(value.substring(1));
	    	    rpr.setColor(color);
	    	} else if (command.startsWith("border")) {
	    		if (ppr == null)
	    			ppr = factory.createPPr();
	    			style.setPPr(ppr);
	    		if (border == null) {
	    			border = factory.createPPrBasePBdr();
	    			ppr.setPBdr(border);
	    		}
	    		boolean[] which = {false, false, false, false};
	    		if (command.equals("border")) {
	    			for (int i=0; i<4; i++) { 
	    				which[i] = true; 
	    			}
	    		} else {
	    			which["trbl".indexOf(command.toLowerCase().charAt(7))] = true;
	    		}
	    		for (int i=0; i<4; i++) {
	    			if (!which[i]) {
	    				continue;
	    			}
	    			CTBorder ctb = factory.createCTBorder();
	    			if (value.equals("none"))
	    				ctb.setVal(STBorder.NONE);
	    			else {
	    				String[] bits = value.split(" ");
	    				if (bits[0].equals("solid")) {
	    					ctb.setVal(STBorder.SINGLE);
	    				} else if (bits[0].equals("dashed")) {
	    					ctb.setVal(STBorder.DASHED);
	    				} else if (bits[0].equals("dotted")) {
	    					ctb.setVal(STBorder.DOTTED);
	    				}
	    				//Color
	    	    	    if (!bits[1].startsWith("#"))
	    	    	    	throw new IllegalArgumentException("Only RGB colors supported for now");
	    	    	    ctb.setColor(bits[1].substring(1));
	    				ctb.setSpace(new BigInteger("0"));
	    				if (bits[2].endsWith("px"))
	    					bits[2] = bits[2].substring(0, bits[2].length()-2);
	    				ctb.setSz(new BigInteger(bits[2]));
	    			}
	    			switch (i) {
	    				case 0:
	    					border.setTop(ctb);
	    					break;
	    				case 1:
	    					border.setRight(ctb);
	    					break;
	    				case 2:
	    					border.setBottom(ctb);
	    					break;
	    				case 3:
	    					border.setLeft(ctb);
	    					break;
	    			}	
	    		}
	    	} else if (command.equals("font-size")) {
				if (value.endsWith("pt"))
					value = value.substring(0, value.length()-2);
    		
				//Word stores in half pts
				value = String.valueOf((int)Math.round(Double.parseDouble(value) * 2));
	    		if (rpr == null) {
	    			rpr = factory.createRPr();
	    			style.setRPr(rpr);
	    		}

	    	    HpsMeasure measure = factory.createHpsMeasure();
	    	    measure.setVal(new BigInteger(value));
	    	    rpr.setSz(measure);
	    	} else if (command.startsWith("margin")) {
	    		if (ppr == null) {
	    			ppr = factory.createPPr();
	    			style.setPPr(ppr);
	    		}
	    		if (spacing == null) {
	    			spacing = factory.createPPrBaseSpacing();
	    			ppr.setSpacing(spacing);
	    		}
	    		if (value.endsWith("px"))
	    			value = value.substring(0, value.length()-2);
	    		if (command.equals("margin-top")) {
	    			spacing.setBefore(new BigInteger(value));
	    		} else if (command.equals("margin-bottom")) {
	    			spacing.setAfter(new BigInteger(value));
	    		}
	    	} else if (command.equals("text-align")) {
	    		if (ppr == null) {
	    			ppr = factory.createPPr();
	    			style.setPPr(ppr);
	    		}
	    		Jc jc = factory.createJc();
	    		jc.setVal(JcEnumeration.valueOf(value.toUpperCase()));
	    		ppr.setJc(jc);
	    	}
	    }
	    
		
	    boolean added = false;
//	    System.out.println(styles);
//	    System.out.println(styles.getJaxbElement());
//	    System.out.println(styles.getJaxbElement().getStyle());
//	    System.out.println(styles.getJaxbElement().getStyle().size());
		try {
			for (int i = 0; i < styles.getContents().getStyle().size(); i++) {
				if (styles.getContents().getStyle().get(i).getName().equals(styleName)) {
					styles.getContents().getStyle().set(i, style);
					added = true;
					break;
				}
			}
			if (!added) {
				styles.getContents().getStyle().add(style);
			}
		} catch (Docx4JException de) {
			de.printStackTrace();
		}
	}
}
