package com.jumbletree.docx5j.docx;

import java.math.BigInteger;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.docx4j.wml.CTShd;
import org.docx4j.wml.CTTabStop;
import org.docx4j.wml.Color;
import org.docx4j.wml.ContentAccessor;
import org.docx4j.wml.P;
import org.docx4j.wml.PPr;
import org.docx4j.wml.PPrBase.PBdr;
import org.docx4j.wml.PPrBase.PStyle;
import org.docx4j.wml.R;
import org.docx4j.wml.STTabJc;
import org.docx4j.wml.STTabTlc;
import org.docx4j.wml.Tabs;
import org.docx4j.wml.Text;

public class Paragraph {

	private DOCXBuilder factory;
	private P p;
	private PPr ppr;
	private Tabs tabs;
	private R run;

	public Paragraph(DOCXBuilder factory) {
		this(factory, null, false);
	}
	
	public Paragraph(DOCXBuilder factory, String style, boolean pageBreakBefore) {
		this.factory = factory;
		this.p = factory.factory.createP();
		if (style != null) {
			PPr ppr = getPPr();
			PStyle pstyle = factory.factory.createPPrBasePStyle();
			pstyle.setVal(style);
			ppr.setPStyle(pstyle);
		}
		if (pageBreakBefore) {
			PPr ppr = getPPr();
			org.docx4j.wml.BooleanDefaultTrue b = factory.factory.createBooleanDefaultTrue();
		    b.setVal(true);	    
			ppr.setPageBreakBefore(b);
		}
	}

	private PPr getPPr() {
		if (this.ppr == null) {
		    PPr ppr = factory.factory.createPPr();	    
		    p.setPPr( ppr );
		    this.ppr = ppr;
		}
		return this.ppr;
	}

	public Paragraph addText(String content) {

		Text text = factory.factory.createText();
		text.setValue(content);

		if (run == null) {
			run = factory.factory.createR();
			p.getContent().add(run);
		}
		run.getContent().add(text);
			
	    return this;
	}
	
	public Paragraph tab() {
		if (run == null) {
			run = factory.factory.createR();
			p.getContent().add(run);
		}
		run.getContent().add(factory.factory.createRTab());
		return this;
	}
	public Paragraph addBoldText(String content) {
		
		Text text = factory.factory.createText();
		text.setValue(content);

		R run = factory.factory.createR();
		run.getContent().add(text);		
		
		p.getContent().add(run);
		
		org.docx4j.wml.RPr rpr = factory.factory.createRPr();		
		org.docx4j.wml.BooleanDefaultTrue b = factory.factory.createBooleanDefaultTrue();
	    b.setVal(true);	    
	    rpr.setB(b);
	    
		run.setRPr(rpr);
		
		// Optionally, set pPr/rPr@w:b		
	    PPr ppr = getPPr();
	    org.docx4j.wml.ParaRPr paraRpr = factory.factory.createParaRPr();
	    ppr.setRPr(paraRpr);	    
	    rpr.setB(b);
		
	    return this;
	}
	
	public Paragraph addItalicText(String content) {
		
		Text text = factory.factory.createText();
		text.setValue(content);

		R run = factory.factory.createR();
		run.getContent().add(text);		
		
		p.getContent().add(run);
		
		org.docx4j.wml.RPr rpr = factory.factory.createRPr();		
		org.docx4j.wml.BooleanDefaultTrue b = factory.factory.createBooleanDefaultTrue();
	    b.setVal(true);	    
	    rpr.setI(b);
	    
		run.setRPr(rpr);
		
		// Optionally, set pPr/rPr@w:i		
	    PPr ppr = getPPr();
	    org.docx4j.wml.ParaRPr paraRpr = factory.factory.createParaRPr();
	    ppr.setRPr(paraRpr);	    
	    rpr.setI(b);
		
	    return this;
	}

	/**
	 * Converts &lt;b&gt; and &lt;i&gt; tags only (at this time).  Doesn't handle nested (ie bold italic) tags...
	 * @param html
	 * @return
	 */
	public Paragraph addFromHTML(String html) {
		Pattern pattern = Pattern.compile("<(span|b|i)(.*?)>(.*?)<\\/\\1>");
		
		Matcher matcher = pattern.matcher(html);

		int start = 0;
		while (matcher.find()) {
			if (matcher.start() > start) {
				addText(html.substring(start, matcher.start()));
			}
			
			String text = matcher.group(3);
			switch (matcher.group(1).toLowerCase().charAt(0)) {
				case 'b':
					addBoldText(text);
					break;
				case 'i':
					addItalicText(text);
					break;
				case 's':
					addStyledText(text, matcher.group(2));
					break;
			}
			
			start = matcher.end();
		}
		
		if (start < html.length())
			addText(html.substring(start));
		
		return this;
	}
	
	private void addStyledText(String content, String style) {
		if (!style.trim().startsWith("style=")) {
			addText(content);
		}
		style = style.trim().substring(6);
		if (style.startsWith("\"")) {
			style = style.substring(1, style.length()-1);
		}
		
		//Create the XML objects
		Text text = factory.factory.createText();
		text.setValue(content);

		R run = factory.factory.createR();
		run.getContent().add(text);		
		
		p.getContent().add(run);
		
		org.docx4j.wml.RPr rpr = factory.factory.createRPr();		
		run.setRPr(rpr);
		
	    getPPr();
////	    org.docx4j.wml.ParaRPr paraRpr = factory.createParaRPr();
////	    ppr.setRPr(paraRpr);	    
		
		for (String bit : style.split(";")) {
			String[] kvp = bit.trim().split(":");
			String key = kvp[0].trim().toLowerCase();
			String val = kvp[1].trim().toLowerCase();
			
			switch (key.charAt(0)) {
				case 'f':
					if (key.equals("font-weight")) {
						if (val.equals("bold")) {
						    org.docx4j.wml.BooleanDefaultTrue b = factory.factory.createBooleanDefaultTrue();
						    b.setVal(true);	    
						    rpr.setB(b);
						}
					} else if (key.equals("font-style")) {
						if (val.equals("italic")) {
						    org.docx4j.wml.BooleanDefaultTrue b = factory.factory.createBooleanDefaultTrue();
						    b.setVal(true);	    
						    rpr.setI(b);
						}
					}
					break;
				case 'c':
					if (key.equals("color")) {
						Color color = factory.factory.createColor();
						if (val.startsWith("#"))
							val = val.substring(1).toUpperCase();
						color.setVal(val);
						rpr.setColor(color);
					}
					break;
				case 'b':
					if (key.equals("background-color")) {
						CTShd shd = factory.factory.createCTShd();
						shd.setFill(val);
						rpr.setShd(shd);
					}
					break;
			}
			
		}
		
	}

	public void addTo(ContentAccessor part) {
		part.getContent().add(p);
	}

	public Paragraph setStyle(String css) {
		for (String rule : css.split(";")) {
			String[] bits = rule.split(":");
			String command = bits[0].trim();
			
			//TODO other rules!
			if (command.equals("border")) {
				String borderCSS = bits[1].trim();
				PBdr borders = factory.factory.createPPrBasePBdr();
				borders.setTop(factory.docx4jCreateBorderDefinition(borderCSS));
				borders.setLeft(factory.docx4jCreateBorderDefinition(borderCSS));
				borders.setRight(factory.docx4jCreateBorderDefinition(borderCSS));
				borders.setBottom(factory.docx4jCreateBorderDefinition(borderCSS));

				ppr.setPBdr(borders);
			} else {
				throw new IllegalArgumentException("Unknown rule: " + rule);
			}
		}
		return this;
	}

	public Paragraph addTabDefinition(int position, String type, String leader) {
		if (this.tabs == null) {
			this.tabs = factory.factory.createTabs();
			ppr.setTabs(tabs);
		}
		CTTabStop tab = factory.factory.createCTTabStop();
		tab.setVal(STTabJc.valueOf(type.toUpperCase()));
		tab.setPos(BigInteger.valueOf(position));
		tab.setLeader(STTabTlc.valueOf(leader.toUpperCase()));
		this.tabs.getTab().add(tab);
		
		return this;
	}
	
	public DOCXBuilder document() {
		return factory;
	}
	
//	public static void main(String[] args) {
//		Pattern pattern = Pattern.compile("<(span|b|i)(.*?)>(.*?)<\\/\\1>");
//
//		String html = "Hello <span style=\"color: #ff0000\">there</span> Bob <b>, nice to sea you</b>";
//		
//		Matcher matcher = pattern.matcher(html);
//		matcher.find();
//		System.out.println(matcher.group(1) + "/" + matcher.group(2) + "/" + matcher.group(3));
//		matcher.find();
//		System.out.println(matcher.group(1) + "/" + matcher.group(2) + "/" + matcher.group(3));
//
//		
//	}
}
