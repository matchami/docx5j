package com.jumbletree.docx5j.xlsx.builders.text;

import java.util.List;

import javax.xml.bind.JAXBElement;

import com.jumbletree.docx5j.xlsx.builders.WorkbookBuilder;

public interface TextModifier {
	public void apply(WorkbookBuilder origin, List<JAXBElement<?>> list);
}