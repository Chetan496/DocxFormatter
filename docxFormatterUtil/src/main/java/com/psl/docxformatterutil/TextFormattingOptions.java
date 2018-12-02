package com.psl.docxformatterutil;


public class TextFormattingOptions {

	
	private boolean bold;
	private boolean italic;
	private boolean underline;
	private String color;
	private boolean isCaps;
	private String fontName;
	private int fontSize;
	
	public TextFormattingOptions(boolean bold, boolean italic, 
			boolean underline, String color, boolean isCaps,
			String fontName,
			int fontSize) {
		
		super();
		this.bold = bold;
		this.italic = italic;
		this.underline = underline;
		this.color = color;
		this.isCaps = isCaps;
		this.fontName = fontName;
		this.fontSize = fontSize;
	}
	

	public String getFontName() {
		return fontName;
	}
	
	public boolean isCaps() {
		return isCaps;
	}
	
	public boolean isBold() {
		return bold;
	}
	public void setBold(boolean bold) {
		this.bold = bold;
	}
	public boolean isItalic() {
		return italic;
	}
	public void setItalic(boolean italic) {
		this.italic = italic;
	}
	public boolean isUnderline() {
		return underline;
	}
	public void setUnderline(boolean underline) {
		this.underline = underline;
	}
	public String getColor() {
		return color;
	}
	public void setColor(String color) {
		this.color = color;
	}


	public int getFontSize() {
		return fontSize;
	}


	public void setFontSize(int fontSize) {
		this.fontSize = fontSize;
	}
	
	
	
	
	
}
