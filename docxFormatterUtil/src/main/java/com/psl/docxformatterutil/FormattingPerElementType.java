package com.psl.docxformatterutil;

import java.util.HashMap;
import java.util.Map.Entry;
import java.util.Set;

public class FormattingPerElementType {

	
	private HashMap<String, TextFormattingOptions> configMap;
	
	
	public FormattingPerElementType() {
		configMap = new HashMap<String, TextFormattingOptions>(10);	
	}
	
	
	public FormattingPerElementType addFormattingOptionsForElementType(final String elementType, 
			final TextFormattingOptions textFormattingOptions) {
		
		configMap.put(elementType, textFormattingOptions);
		
		return this;
	}
	
	
	public TextFormattingOptions getFormattingOptions(final String elementType) {
		
		return configMap.get(elementType);
		
	}
	
	public Set<String> getAllElementTypes(){
		
		return configMap.keySet() ;
	}
	
	
	public Set<Entry<String, TextFormattingOptions>> getEntries(){
		
		return configMap.entrySet();
		
	}
	
	
}
