package com.psl.docxformatterutil;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.math.BigInteger;
import java.util.Arrays;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;
import java.util.Set;
import java.util.logging.Logger;

import javax.xml.bind.JAXBElement;
import javax.xml.bind.JAXBException;

import org.docx4j.dml.BaseStyles.FontScheme;
import org.docx4j.dml.TextFont;
import org.docx4j.dml.Theme;
import org.docx4j.jaxb.Context;
import org.docx4j.jaxb.XPathBinderAssociationIsPartialException;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.ThemePart;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.openpackaging.parts.WordprocessingML.StyleDefinitionsPart;
import org.docx4j.wml.BooleanDefaultTrue;
import org.docx4j.wml.Color;
import org.docx4j.wml.HpsMeasure;
import org.docx4j.wml.ObjectFactory;
import org.docx4j.wml.P;
import org.docx4j.wml.R;
import org.docx4j.wml.RFonts;
import org.docx4j.wml.RPr;
import org.docx4j.wml.Style;
import org.docx4j.wml.Styles;
import org.docx4j.wml.Text;

public class MainDocxFormatter {

	private static Logger logger = Logger.getLogger(MainDocxFormatter.class.getName());

	private static final List<String> SUPPORTED_GLOBAL_STYLES = Arrays.asList("Normal", "Heading1", "Heading2",
			"Title",  "majorFont", "minorFont");

	private static void exampleToFormatDocumentGlobalStylesApproach() {

		FormattingPerElementType formattingPerElemType = new FormattingPerElementType();

		TextFormattingOptions formattingOptionsTitle = new TextFormattingOptions(true, true, false, "orange", true,
				"Courier New", 36);
		formattingPerElemType.addFormattingOptionsForElementType("Title", formattingOptionsTitle);

		TextFormattingOptions formattingOptionsH1 = new TextFormattingOptions(true, false, false, "black", true,
				"Verdana", 28);
		formattingPerElemType.addFormattingOptionsForElementType("Heading1", formattingOptionsH1);

		TextFormattingOptions formattingOptionsH2 = new TextFormattingOptions(true, false, false, "green", true,
				"Times New Roman", 20);
		formattingPerElemType.addFormattingOptionsForElementType("Heading2", formattingOptionsH2);

		TextFormattingOptions formattingOptionsNormal = new TextFormattingOptions(true, false, false, "red", true,
				"Comic Sans MS", 14);
		formattingPerElemType.addFormattingOptionsForElementType("Normal", formattingOptionsNormal);

		TextFormattingOptions formattingOptionsMajorFont = new TextFormattingOptions(true, false, false, "red", true,
				"Nyala", 14);
		formattingPerElemType.addFormattingOptionsForElementType("majorFont", formattingOptionsMajorFont);
		
		TextFormattingOptions formattingOptionsMinorFont = new TextFormattingOptions(true, false, false, "red", true,
				"Arial", 14);
		formattingPerElemType.addFormattingOptionsForElementType("minorFont", formattingOptionsMinorFont);

		
		ObjectFactory factory = Context.getWmlObjectFactory();

		WordprocessingMLPackage wordMLPackage = loadWordDoc("E:\\\\SampleWordResume.docx");
		modifyGlobalStyles(factory, wordMLPackage, formattingPerElemType);

		File exportFile = new File("E:\\modifiedResume.docx");
		try {
			wordMLPackage.save(exportFile);
		} catch (Docx4JException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

	}

	private static void exampleToFormatDocumentXPathBasedApproach() {

		FormattingPerElementType formattingPerElemType = new FormattingPerElementType();

		TextFormattingOptions formattingOptionsH1 = new TextFormattingOptions(true, false, false, "black", true,
				"Calibri", 40);
		formattingPerElemType.addFormattingOptionsForElementType("heading 1", formattingOptionsH1);

		TextFormattingOptions formattingOptionsH2 = new TextFormattingOptions(true, false, false, "green", true,
				"Courier New", 30);
		formattingPerElemType.addFormattingOptionsForElementType("heading 2", formattingOptionsH2);

		TextFormattingOptions formattingOptionsTitle = new TextFormattingOptions(true, false, false, "red", true,
				"Courier New", 30);
		formattingPerElemType.addFormattingOptionsForElementType("Title", formattingOptionsTitle);

		TextFormattingOptions formattingOptionsNormal = new TextFormattingOptions(true, false, false, "red", true,
				"Comic Sans MS", 20);
		formattingPerElemType.addFormattingOptionsForElementType("Normal", formattingOptionsNormal);

		ObjectFactory factory = Context.getWmlObjectFactory();

		WordprocessingMLPackage wordMLPackage = loadWordDoc("E:\\\\SampleWordResume.docx");
		Set<Entry<String, TextFormattingOptions>> entries = formattingPerElemType.getEntries();

		for (final Map.Entry<String, TextFormattingOptions> entry : entries) {
			wordMLPackage = applyFormattingOptions(factory, wordMLPackage, entry);
		}

		File exportFile = new File("E:\\modifiedResume.docx");
		try {
			wordMLPackage.save(exportFile);
		} catch (Docx4JException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

	}

	public static ByteArrayOutputStream FormatWordDoc(final FormattingPerElementType formattingPerElemType,
			final FileInputStream fileInputStream) {

		ObjectFactory factory = Context.getWmlObjectFactory();

		WordprocessingMLPackage wordMLPackage = loadWordDocFromFileInputStream(fileInputStream);
		wordMLPackage = modifyGlobalStyles(factory, wordMLPackage, formattingPerElemType);

		ByteArrayOutputStream byteArrayOutputStream = new ByteArrayOutputStream();
		try {
			wordMLPackage.save(byteArrayOutputStream);
			return byteArrayOutputStream;

		} catch (Docx4JException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

		return null;

	}

	private static WordprocessingMLPackage modifyGlobalStyles(final ObjectFactory factory,
			WordprocessingMLPackage wordMLPackage, final FormattingPerElementType formattingPerElemType) {

		MainDocumentPart mainDocumentPart = wordMLPackage.getMainDocumentPart();
		StyleDefinitionsPart styleDefinitionsPart = mainDocumentPart.getStyleDefinitionsPart();

		// we are modifying the global styles and optionally theme level styles
		Map<String, Style> stringToStyleMap = StyleDefinitionsPart.getKnownStyles();
		Set<Entry<String, TextFormattingOptions>> entries = formattingPerElemType.getEntries();

		ThemePart themePart = mainDocumentPart.getThemePart();
		Theme theme = null;
		try {
			theme = themePart.getContents();
			//logger.info("Theme name is "+theme.getName());
			
		} catch (Docx4JException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

		for (Entry<String, TextFormattingOptions> configEntry : entries) {

			final String configEntryKey = configEntry.getKey();
			// entry from config must be in allowed styles...otherwise its garbage input
			if (SUPPORTED_GLOBAL_STYLES.contains(configEntryKey)) {
				
				final TextFormattingOptions formattingOptions = configEntry.getValue();
				
				if (stringToStyleMap.containsKey(configEntryKey)) {

					// standard style supported by MS-Office
					Style s = styleDefinitionsPart.getStyleById(stringToStyleMap.get(configEntryKey).getStyleId());
					

					RPr rpr = s.getRPr();

					if (rpr == null) {
						rpr = factory.createRPr();
						s.setRPr(rpr);
					}

					RFonts rf = rpr.getRFonts();

					if (rf == null) {

						rf = factory.createRFonts();
						rpr.setRFonts(rf);

					}

					rf.setAscii(formattingOptions.getFontName()); // wont work for headings..
					BooleanDefaultTrue booleanDefaultTrue = new BooleanDefaultTrue();

					if (formattingOptions.isBold()) {
						rpr.setB(booleanDefaultTrue);
					}

					if (formattingOptions.isCaps()) {
						rpr.setCaps(booleanDefaultTrue);
					}

					if (formattingOptions.isItalic()) {
						rpr.setI(booleanDefaultTrue);
					}

					HpsMeasure hpsMeasure = new HpsMeasure();
					hpsMeasure.setVal(new BigInteger("" + (formattingOptions.getFontSize() * 2)));
					rpr.setSz(hpsMeasure);

					Color color = factory.createColor();
					color.setVal(formattingOptions.getColor());

					rpr.setColor(color);

				} else {

					// must be at theme level..need to modify the theme.
					//logger.info("theme level setting.." + configEntryKey);
					FontScheme fontScheme = theme.getThemeElements().getFontScheme();
					if(configEntryKey.equals("majorFont")) {
						
						TextFont majorTextFont =  fontScheme.getMajorFont().getLatin() ;
						majorTextFont.setTypeface(formattingOptions.getFontName());
						
					}else if(configEntryKey.equals("minorFont")) {
						
						TextFont minorTextFont = fontScheme.getMinorFont().getLatin() ; 
						minorTextFont.setTypeface(formattingOptions.getFontName());
					}
					

				}

			} else {
				//logger.info("invalid input");
				continue; // skip this..user has passed something which is not supported
			}
		}

		return wordMLPackage;

	}
	
	

	private static WordprocessingMLPackage applyDefaultStylingToRuns(final ObjectFactory factory,
			WordprocessingMLPackage wordMLPackage, final Map.Entry<String, TextFormattingOptions> formattingEntry) {

		MainDocumentPart mainDocumentPart = wordMLPackage.getMainDocumentPart();

		try {
			Styles styles = mainDocumentPart.getStyleDefinitionsPart().getContents();

			TextFormattingOptions formattingOptions = formattingEntry.getValue();
			final String key = formattingEntry.getKey(); // key is the aspect for which default style to be applied

			if (!SUPPORTED_GLOBAL_STYLES.contains(key)) {
				return wordMLPackage;
			}

			// we are modifying the global styles
			for (final Style s : styles.getStyle()) {

				if (s.getName().getVal().equals(key)) {
					RPr rpr = s.getRPr();

					if (rpr == null) {
						rpr = factory.createRPr();
						s.setRPr(rpr);
					}

					RFonts rf = rpr.getRFonts();

					if (rf == null) {

						rf = factory.createRFonts();
						rpr.setRFonts(rf);

					}
					rf.setAscii(formattingOptions.getFontName());
					BooleanDefaultTrue booleanDefaultTrue = new BooleanDefaultTrue();

					if (formattingOptions.isBold()) {
						rpr.setB(booleanDefaultTrue);
					}

					if (formattingOptions.isCaps()) {
						rpr.setCaps(booleanDefaultTrue);
					}

					if (formattingOptions.isItalic()) {
						rpr.setI(booleanDefaultTrue);
					}

					Color color = factory.createColor();
					color.setVal(formattingOptions.getColor());

					rpr.setColor(color);

				}

			}

		} catch (Docx4JException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

		return wordMLPackage;

	}

	private static WordprocessingMLPackage applyFormattingOptions(final ObjectFactory factory,
			WordprocessingMLPackage wordMLPackage, final Map.Entry<String, TextFormattingOptions> formattingEntry) {

		logger.info("applying formatting for " + formattingEntry.getKey());

		if (formattingEntry.getKey().equals("Normal")) {
			return applyDefaultStylingToRuns(factory, wordMLPackage, formattingEntry);
		}

		MainDocumentPart mainDocumentPart = wordMLPackage.getMainDocumentPart();

		final String xPath = generateXPathForElemType(formattingEntry.getKey());

		logger.info("generated xpath " + xPath);

		try {

			// the false param is required due to the way we are doing this logic
			List<Object> paraNodes = mainDocumentPart.getJAXBNodesViaXPath(xPath, false);

			logger.info("paraNodes is " + paraNodes);

			for (Object obj : paraNodes) {
				P para = (P) obj;

				if (para.getContent() == null || para.getContent().size() == 0) {
					continue;
				}

				Object object = para.getContent().get(0);
				R run;
				if (object instanceof R) {
					run = (R) para.getContent().get(0);
				} else {
					continue;
				}

				RPr rpr = run.getRPr();

				if (rpr == null) {
					rpr = factory.createRPr();
				}

				BooleanDefaultTrue booleanDefaultTrue = new BooleanDefaultTrue();
				TextFormattingOptions textFormattingOptions = formattingEntry.getValue();

				if (textFormattingOptions.isBold()) {
					rpr.setB(booleanDefaultTrue);
				}

				if (textFormattingOptions.isCaps()) {
					rpr.setCaps(booleanDefaultTrue);
				}

				if (textFormattingOptions.isItalic()) {
					rpr.setI(booleanDefaultTrue);
				}

				RFonts rfonts = factory.createRFonts();
				rfonts.setAscii(textFormattingOptions.getFontName());

				rpr.setRFonts(rfonts);

				Color color = factory.createColor();
				color.setVal(textFormattingOptions.getColor());

				rpr.setColor(color);
				run.setRPr(rpr);

			}

		} catch (XPathBinderAssociationIsPartialException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (JAXBException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} finally {

		}

		return wordMLPackage;

	}

	private static String generateXPathForElemType(final String elementType) {

		if (elementType.equals("heading 1")) {
			return "//w:p[w:pPr[w:pStyle[@w:val='Heading1']]]";
		}

		if (elementType.equals("heading 2")) {
			return "//w:p[w:pPr[w:pStyle[@w:val='Heading2']]]";

		}
		if (elementType.equals("heading 3")) {
			return "//w:p[w:pPr[w:pStyle[@w:val='Heading3']]]";

		}

		if (elementType.equals("Title")) {
			return "//w:p[w:pPr[w:pStyle[@w:val='Title']]]";

		}

		return null;

	}

	private static void readTextFromWordDoc(final String filePath) {

		File doc = new File(filePath);
		WordprocessingMLPackage wordMLPackage;
		try {
			wordMLPackage = WordprocessingMLPackage.load(doc);
			MainDocumentPart mainDocumentPart = wordMLPackage.getMainDocumentPart();

			String textNodesXPath = "//w:t";

			List<Object> textNodes = mainDocumentPart.getJAXBNodesViaXPath(textNodesXPath, true);

			for (Object obj : textNodes) {
				Text text = (Text) ((JAXBElement) obj).getValue();
				String textValue = text.getValue();
				System.out.println(textValue);
			}

		} catch (Docx4JException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (JAXBException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

	}

	private static WordprocessingMLPackage loadWordDoc(final String wordDocPath) {

		File doc = new File(wordDocPath);
		try {
			WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.load(doc);
			return wordMLPackage;
		} catch (Docx4JException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		;

		return null;

	}

	private static WordprocessingMLPackage loadWordDocFromFileInputStream(final FileInputStream fileInputStream) {

		try {
			WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.load(fileInputStream);
			return wordMLPackage;
		} catch (Docx4JException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		;

		return null;

	}

}
