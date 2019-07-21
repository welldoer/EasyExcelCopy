package net.blogjava.easyexcelcopy.spike;

import static org.assertj.core.api.Assertions.*;

import java.io.ByteArrayInputStream;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;

import javax.xml.parsers.ParserConfigurationException;
import javax.xml.parsers.SAXParserFactory;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellReference;
import org.junit.Before;
import org.junit.Test;
import org.xml.sax.Attributes;
import org.xml.sax.InputSource;
import org.xml.sax.Locator;
import org.xml.sax.SAXException;
import org.xml.sax.SAXParseException;
import org.xml.sax.XMLReader;
import org.xml.sax.helpers.DefaultHandler;

public class SampleExcelXMLTest {
	private XMLReader xmlReader;
	
	private String sharedStringsXml;
	private String workbookXml;
	private String sheet1Xml;
	private String sheet2Xml;

	@Before
	public void setUp() throws Exception {
		xmlReader = SAXParserFactory.newInstance().newSAXParser().getXMLReader();

		/*
		 * <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
		 *   <sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="51" uniqueCount="14">
		 * 	 <si><t>表格 1</t></si>
		 * 	 <si><t>银行放款编号</t></si>
		 *   <si><t>C0de</t></si>
		 *   <si><t>银行放款日期</t></si>
		 *   <si><t>银行放款金额</t></si>
		 *   <si><t>银行利率</t></si>
		 *   <si><t>银行借款期限</t></si>
		 *   <si><t>到期日</t></si>
		 *   <si><t>每月应还利息</t></si>
		 *   <si><t>借款人名称</t></si>
		 *   <si><t>深圳市祥合鑫科技有限公司</t></si>
		 *   <si><t>深圳市迈赛特光电有限公司</t></si>
		 *   <si><t>流水贷上海系统第431214号</t></si>
		 *   <si><t>青岛博纳服饰有限公司</t></si>
		 * </sst>
		 */
		sharedStringsXml = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n" + 
				"<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"51\" uniqueCount=\"14\"><si><t>表格 1</t></si><si><t>银行放款编号</t></si><si><t>C0de</t></si><si><t>银行放款日期</t></si><si><t>银行放款金额</t></si><si><t>银行利率</t></si><si><t>银行借款期限</t></si><si><t>到期日</t></si><si><t>每月应还利息</t></si><si><t>借款人名称</t></si><si><t>深圳市祥合鑫科技有限公司</t></si><si><t>深圳市迈赛特光电有限公司</t></si><si><t>流水贷上海系统第431214号</t></si><si><t>青岛博纳服饰有限公司</t></si></sst>";

		/*
		 * <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
		 *   <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
		 *     <fileVersion appName="xl" lastEdited="5" lowestEdited="5" rupBuild="25915"/>
		 *     <workbookPr date1904="1" showInkAnnotation="0" autoCompressPictures="0"/>
		 *     <bookViews>
		 *       <workbookView xWindow="33740" yWindow="320" windowWidth="20340" windowHeight="18080" activeTab="1"/>
		 *     </bookViews>
		 *     <sheets>
		 *       <sheet name="工作表 1" sheetId="1" r:id="rId1"/>
		 *       <sheet name="工作表1" sheetId="2" r:id="rId2"/>
		 *     </sheets>
		 *     <calcPr calcId="4294967295" concurrentCalc="0"/>
		 *     <extLst>
		 *       <ext uri="{7523E5D3-25F3-A5E0-1632-64F254C22452}" xmlns:mx="http://schemas.microsoft.com/office/mac/excel/2008/main">
		 *         <mx:ArchID Flags="2"/>
		 *       </ext>
		 *     </extLst>
		 *   </workbook>
		 */
		workbookXml = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n" + 
				"<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><fileVersion appName=\"xl\" lastEdited=\"5\" lowestEdited=\"5\" rupBuild=\"25915\"/><workbookPr date1904=\"1\" showInkAnnotation=\"0\" autoCompressPictures=\"0\"/><bookViews><workbookView xWindow=\"33740\" yWindow=\"320\" windowWidth=\"20340\" windowHeight=\"18080\" activeTab=\"1\"/></bookViews><sheets><sheet name=\"工作表 1\" sheetId=\"1\" r:id=\"rId1\"/><sheet name=\"工作表1\" sheetId=\"2\" r:id=\"rId2\"/></sheets><calcPr calcId=\"4294967295\" concurrentCalc=\"0\"/><extLst><ext uri=\"{7523E5D3-25F3-A5E0-1632-64F254C22452}\" xmlns:mx=\"http://schemas.microsoft.com/office/mac/excel/2008/main\"><mx:ArchID Flags=\"2\"/></ext></extLst></workbook>";

		/*
		 * <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
		 * <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">
		 *   <sheetPr enableFormatConditionsCalculation="0">
		 *     <pageSetUpPr fitToPage="1"/>
		 *   </sheetPr>
		 *   <dimension ref="A1:I23"/>
		 *   <sheetViews>
		 *     <sheetView showGridLines="0" workbookViewId="0">
		 *       <selection activeCell="A2" sqref="A2:XFD18"/>
		 *     </sheetView>
		 *   </sheetViews>
		 *   <sheetFormatPr baseColWidth="10" defaultColWidth="16.83203125" defaultRowHeight="18" customHeight="1" x14ac:dyDescent="0"/>
		 *   <cols>
		 *     <col min="1" max="9" width="16.33203125" style="1" customWidth="1"/>
		 *     <col min="10" max="16384" width="16.83203125" style="1"/>
		 *   </cols>
		 *   <sheetData>
		 *     <row r="1" spans="1:9" ht="31" customHeight="1">
		 *       <c r="A1" s="34" t="s"><v>0</v></c>
		 *       <c r="B1" s="35"/>
		 *       <c r="C1" s="36"/>
		 *       <c r="D1" s="36"/>
		 *       <c r="E1" s="35"/>
		 *       <c r="F1" s="35"/>
		 *       <c r="G1" s="35"/>
		 *       <c r="H1" s="35"/>
		 *       <c r="I1" s="37"/>
		 *     </row>
		 *     <row r="2" spans="1:9" ht="23" customHeight="1">
		 *       <c r="A2" s="2" t="s"><v>1</v></c>
		 *       <c r="B2" s="3" t="s"><v>2</v></c>
		 *       <c r="C2" s="4" t="s"><v>3</v></c>
		 *       <c r="D2" s="5" t="s"><v>4</v></c>
		 *       <c r="E2" s="6" t="s"><v>5</v></c>
		 *       <c r="F2" s="7" t="s"><v>6</v></c>
		 *       <c r="G2" s="8" t="s"><v>7</v></c>
		 *       <c r="H2" s="2" t="s"><v>8</v></c>
		 *       <c r="I2" s="4" t="s"><v>9</v></c>
		 *     </row>
		 *     <row r="3" spans="1:9" ht="21" customHeight="1">
		 *       <c r="A3" s="9"><v>1</v></c>
		 *       <c r="B3" s="10"><v>431212</v></c>
		 *       <c r="C3" s="11"><v>41242</v></c>
		 *       <c r="D3" s="12"><v>11</v></c>
		 *       <c r="E3" s="13"><v>0.05</v></c>
		 *       <c r="F3" s="14"><v>1</v></c>
		 *       <c r="G3" s="15"><v>41242</v></c>
		 *       <c r="H3" s="16"><v>12</v></c>
		 *       <c r="I3" s="17" t="s"><v>10</v></c>
		 *     </row>
		 *     <row r="4" spans="1:9" ht="21" customHeight="1">
		 *       <c r="A4" s="18"><v>2</v></c>
		 *       <c r="B4" s="19"><v>6666</v></c>
		 *       <c r="C4" s="20"><v>41243</v></c>
		 *       <c r="D4" s="21"><v>753578.16599999997</v></c>
		 *       <c r="E4" s="22"><v>0.05</v></c>
		 *       <c r="F4" s="23"><v>1</v></c>
		 *       <c r="G4" s="15"><v>41243</v></c>
		 *       <c r="H4" s="24"><v>12</v></c>
		 *       <c r="I4" s="25" t="s"><v>11</v></c>
		 *     </row>
		 *     <row r="5" spans="1:9" ht="21" customHeight="1">
		 *       <c r="A5" s="18"><v>3</v></c>
		 *       <c r="B5" s="19"><v>431214</v></c>
		 *       <c r="C5" s="20"><v>41244</v></c>
		 *       <c r="D5" s="21"><v>3452</v></c>
		 *       <c r="E5" s="22"><v>0.05</v></c>
		 *       <c r="F5" s="23"><v>2</v></c>
		 *       <c r="G5" s="15"><v>41244</v></c>
		 *       <c r="H5" s="24"><v>12</v></c>
		 *       <c r="I5" s="25" t="s"><v>12</v></c>
		 *     </row>
		 *     <row r="6" spans="1:9" ht="21" customHeight="1">
		 *       <c r="A6" s="18"><v>4</v></c>
		 *       <c r="B6" s="19"><v>431214</v></c>
		 *       <c r="C6" s="20"><v>41247</v></c>
		 *       <c r="D6" s="21"><v>8567</v></c>
		 *       <c r="E6" s="22"><v>0.05</v></c>
		 *       <c r="F6" s="23"><v>6</v></c>
		 *       <c r="G6" s="15"><v>41247</v></c>
		 *       <c r="H6" s="24"><v>12</v></c>
		 *       <c r="I6" s="25" t="s"><v>12</v></c>
		 *     </row><row r="7" spans="1:9" ht="20.5" customHeight="1"><c r="A7" s="18"><v>5</v></c><c r="B7" s="19"><v>431214</v></c><c r="C7" s="20"><v>41247</v></c><c r="D7" s="21"><v>7123</v></c><c r="E7" s="22"><v>0.05</v></c><c r="F7" s="23"><v>13</v></c><c r="G7" s="15"><v>41247</v></c><c r="H7" s="24"><v>12</v></c><c r="I7" s="25" t="s"><v>12</v></c></row><row r="8" spans="1:9" ht="20.25" customHeight="1"><c r="A8" s="18"><v>6</v></c><c r="B8" s="19"><v>431214</v></c><c r="C8" s="20"><v>41247</v></c><c r="D8" s="21"><v>6000</v></c><c r="E8" s="22"><v>0.05</v></c><c r="F8" s="23"><v>4</v></c><c r="G8" s="15"><v>41247</v></c><c r="H8" s="24"><v>12</v></c><c r="I8" s="25" t="s"><v>12</v></c></row><row r="9" spans="1:9" ht="20.25" customHeight="1"><c r="A9" s="18"><v>7</v></c><c r="B9" s="19"><v>431214</v></c><c r="C9" s="20"><v>41247</v></c><c r="D9" s="21"><v>50000</v></c><c r="E9" s="22"><v>0.05</v></c><c r="F9" s="23"><v>3</v></c><c r="G9" s="15"><v>41247</v></c><c r="H9" s="24"><v>12</v></c><c r="I9" s="25" t="s"><v>12</v></c></row><row r="10" spans="1:9" ht="20.25" customHeight="1"><c r="A10" s="18"><v>8</v></c><c r="B10" s="19"><v>431214</v></c><c r="C10" s="20"><v>41247</v></c><c r="D10" s="21"><v>120000</v></c><c r="E10" s="22"><v>0.05</v></c><c r="F10" s="23"><v>1</v></c><c r="G10" s="15"><v>41247</v></c><c r="H10" s="24"><v>12</v></c><c r="I10" s="25" t="s"><v>12</v></c></row><row r="11" spans="1:9" ht="20.25" customHeight="1"><c r="A11" s="18"><v>9</v></c><c r="B11" s="26"><v>3695064866</v></c><c r="C11" s="27"><v>41247</v></c><c r="D11" s="21"><v>1234</v></c><c r="E11" s="22"><v>0.05</v></c><c r="F11" s="23"><v>2</v></c><c r="G11" s="15"><v>41247</v></c><c r="H11" s="24"><v>12</v></c><c r="I11" s="25" t="s"><v>13</v></c></row><row r="12" spans="1:9" ht="20.25" customHeight="1"><c r="A12" s="18"><v>10</v></c><c r="B12" s="26"><v>3695064866</v></c><c r="C12" s="27"><v>41247</v></c><c r="D12" s="21"><v>122</v></c><c r="E12" s="22"><v>0.05</v></c><c r="F12" s="23"><v>1</v></c><c r="G12" s="15"><v>41247</v></c><c r="H12" s="24"><v>12</v></c><c r="I12" s="25" t="s"><v>13</v></c></row><row r="13" spans="1:9" ht="20.25" customHeight="1"><c r="A13" s="18"><v>11</v></c><c r="B13" s="26"><v>3695064866</v></c><c r="C13" s="27"><v>41247</v></c><c r="D13" s="21"><v>781115.19</v></c><c r="E13" s="22"><v>0.05</v></c><c r="F13" s="23"><v>1</v></c><c r="G13" s="15"><v>41247</v></c><c r="H13" s="24"><v>12</v></c><c r="I13" s="25" t="s"><v>13</v></c></row><row r="14" spans="1:9" ht="20.25" customHeight="1"><c r="A14" s="18"><v>12</v></c><c r="B14" s="26"><v>3695064866</v></c><c r="C14" s="27"><v>41247</v></c><c r="D14" s="21"><v>781115.19</v></c><c r="E14" s="22"><v>0.05</v></c><c r="F14" s="23"><v>1</v></c><c r="G14" s="15"><v>41247</v></c><c r="H14" s="24"><v>12</v></c><c r="I14" s="25" t="s"><v>13</v></c></row><row r="15" spans="1:9" ht="20.25" customHeight="1"><c r="A15" s="18"><v>13</v></c><c r="B15" s="26"><v>3695064866</v></c><c r="C15" s="27"><v>41247</v></c><c r="D15" s="21"><v>781115.19</v></c><c r="E15" s="22"><v>0.05</v></c><c r="F15" s="23"><v>1</v></c><c r="G15" s="15"><v>41247</v></c><c r="H15" s="24"><v>12</v></c><c r="I15" s="25" t="s"><v>13</v></c></row><row r="16" spans="1:9" ht="20.25" customHeight="1"><c r="A16" s="18"><v>14</v></c><c r="B16" s="26"><v>3695064866</v></c><c r="C16" s="27"><v>41247</v></c><c r="D16" s="21"><v>781115.19</v></c><c r="E16" s="22"><v>0.05</v></c><c r="F16" s="23"><v>1</v></c><c r="G16" s="15"><v>41247</v></c><c r="H16" s="24"><v>12</v></c><c r="I16" s="25" t="s"><v>13</v></c></row><row r="17" spans="1:9" ht="20.25" customHeight="1"><c r="A17" s="18"><v>15</v></c><c r="B17" s="26"><v>3695064866</v></c><c r="C17" s="27"><v>41247</v></c><c r="D17" s="21"><v>781115.19</v></c><c r="E17" s="22"><v>0.05</v></c><c r="F17" s="23"><v>1</v></c><c r="G17" s="15"><v>41247</v></c><c r="H17" s="24"><v>12</v></c><c r="I17" s="25" t="s"><v>13</v></c></row><row r="18" spans="1:9" ht="20.25" customHeight="1"><c r="A18" s="18"><v>16</v></c><c r="B18" s="26"><v>3695064866</v></c><c r="C18" s="28"><v>41247</v></c><c r="D18" s="29"><v>781115.19</v></c><c r="E18" s="22"><v>0.05</v></c><c r="F18" s="23"><v>1</v></c><c r="G18" s="30"><v>41247</v></c><c r="H18" s="24"><v>12</v></c><c r="I18" s="31" t="s"><v>13</v></c></row><row r="19" spans="1:9" ht="20.25" customHeight="1"><c r="A19" s="18"/><c r="B19" s="32"/><c r="C19" s="33"/><c r="D19" s="33"/><c r="E19" s="33"/><c r="F19" s="33"/><c r="G19" s="33"/><c r="H19" s="33"/><c r="I19" s="33"/></row><row r="20" spans="1:9" ht="20.25" customHeight="1"><c r="A20" s="18"/><c r="B20" s="32"/><c r="C20" s="33"/><c r="D20" s="33"/><c r="E20" s="33"/><c r="F20" s="33"/><c r="G20" s="33"/><c r="H20" s="33"/><c r="I20" s="33"/></row><row r="21" spans="1:9" ht="20.25" customHeight="1"><c r="A21" s="18"/><c r="B21" s="32"/><c r="C21" s="33"/><c r="D21" s="33"/><c r="E21" s="33"/><c r="F21" s="33"/><c r="G21" s="33"/><c r="H21" s="33"/><c r="I21" s="33"/></row><row r="22" spans="1:9" ht="20.25" customHeight="1"><c r="A22" s="18"/><c r="B22" s="32"/><c r="C22" s="33"/><c r="D22" s="33"/><c r="E22" s="33"/><c r="F22" s="33"/><c r="G22" s="33"/><c r="H22" s="33"/><c r="I22" s="33"/></row><row r="23" spans="1:9" ht="20.25" customHeight="1"><c r="A23" s="18"/><c r="B23" s="32"/><c r="C23" s="33"/><c r="D23" s="33"/><c r="E23" s="33"/><c r="F23" s="33"/><c r="G23" s="33"/><c r="H23" s="33"/><c r="I23" s="33"/></row></sheetData><mergeCells count="1"><mergeCell ref="A1:I1"/></mergeCells><phoneticPr fontId="5" type="noConversion"/><pageMargins left="0.5" right="0.5" top="0.75" bottom="0.75" header="0.27777779102325439" footer="0.27777779102325439"/><pageSetup paperSize="9" orientation="portrait" horizontalDpi="4294967292" verticalDpi="4294967292"/><headerFooter><oddFooter>&amp;C&amp;"Helvetica,Regular"&amp;12&amp;K000000&amp;P</oddFooter></headerFooter><extLst><ext uri="{64002731-A6B0-56B0-2670-7721B7C09600}" xmlns:mx="http://schemas.microsoft.com/office/mac/excel/2008/main"><mx:PLV Mode="0" OnePage="0" WScale="0"/></ext></extLst></worksheet>
		 */
		sheet1Xml = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n" + 
				"<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" mc:Ignorable=\"x14ac\" xmlns:x14ac=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac\"><sheetPr enableFormatConditionsCalculation=\"0\"><pageSetUpPr fitToPage=\"1\"/></sheetPr><dimension ref=\"A1:I23\"/><sheetViews><sheetView showGridLines=\"0\" workbookViewId=\"0\"><selection activeCell=\"A2\" sqref=\"A2:XFD18\"/></sheetView></sheetViews><sheetFormatPr baseColWidth=\"10\" defaultColWidth=\"16.83203125\" defaultRowHeight=\"18\" customHeight=\"1\" x14ac:dyDescent=\"0\"/><cols><col min=\"1\" max=\"9\" width=\"16.33203125\" style=\"1\" customWidth=\"1\"/><col min=\"10\" max=\"16384\" width=\"16.83203125\" style=\"1\"/></cols><sheetData><row r=\"1\" spans=\"1:9\" ht=\"31\" customHeight=\"1\"><c r=\"A1\" s=\"34\" t=\"s\"><v>0</v></c><c r=\"B1\" s=\"35\"/><c r=\"C1\" s=\"36\"/><c r=\"D1\" s=\"36\"/><c r=\"E1\" s=\"35\"/><c r=\"F1\" s=\"35\"/><c r=\"G1\" s=\"35\"/><c r=\"H1\" s=\"35\"/><c r=\"I1\" s=\"37\"/></row><row r=\"2\" spans=\"1:9\" ht=\"23\" customHeight=\"1\"><c r=\"A2\" s=\"2\" t=\"s\"><v>1</v></c><c r=\"B2\" s=\"3\" t=\"s\"><v>2</v></c><c r=\"C2\" s=\"4\" t=\"s\"><v>3</v></c><c r=\"D2\" s=\"5\" t=\"s\"><v>4</v></c><c r=\"E2\" s=\"6\" t=\"s\"><v>5</v></c><c r=\"F2\" s=\"7\" t=\"s\"><v>6</v></c><c r=\"G2\" s=\"8\" t=\"s\"><v>7</v></c><c r=\"H2\" s=\"2\" t=\"s\"><v>8</v></c><c r=\"I2\" s=\"4\" t=\"s\"><v>9</v></c></row><row r=\"3\" spans=\"1:9\" ht=\"21\" customHeight=\"1\"><c r=\"A3\" s=\"9\"><v>1</v></c><c r=\"B3\" s=\"10\"><v>431212</v></c><c r=\"C3\" s=\"11\"><v>41242</v></c><c r=\"D3\" s=\"12\"><v>11</v></c><c r=\"E3\" s=\"13\"><v>0.05</v></c><c r=\"F3\" s=\"14\"><v>1</v></c><c r=\"G3\" s=\"15\"><v>41242</v></c><c r=\"H3\" s=\"16\"><v>12</v></c><c r=\"I3\" s=\"17\" t=\"s\"><v>10</v></c></row><row r=\"4\" spans=\"1:9\" ht=\"21\" customHeight=\"1\"><c r=\"A4\" s=\"18\"><v>2</v></c><c r=\"B4\" s=\"19\"><v>6666</v></c><c r=\"C4\" s=\"20\"><v>41243</v></c><c r=\"D4\" s=\"21\"><v>753578.16599999997</v></c><c r=\"E4\" s=\"22\"><v>0.05</v></c><c r=\"F4\" s=\"23\"><v>1</v></c><c r=\"G4\" s=\"15\"><v>41243</v></c><c r=\"H4\" s=\"24\"><v>12</v></c><c r=\"I4\" s=\"25\" t=\"s\"><v>11</v></c></row><row r=\"5\" spans=\"1:9\" ht=\"21\" customHeight=\"1\"><c r=\"A5\" s=\"18\"><v>3</v></c><c r=\"B5\" s=\"19\"><v>431214</v></c><c r=\"C5\" s=\"20\"><v>41244</v></c><c r=\"D5\" s=\"21\"><v>3452</v></c><c r=\"E5\" s=\"22\"><v>0.05</v></c><c r=\"F5\" s=\"23\"><v>2</v></c><c r=\"G5\" s=\"15\"><v>41244</v></c><c r=\"H5\" s=\"24\"><v>12</v></c><c r=\"I5\" s=\"25\" t=\"s\"><v>12</v></c></row><row r=\"6\" spans=\"1:9\" ht=\"21\" customHeight=\"1\"><c r=\"A6\" s=\"18\"><v>4</v></c><c r=\"B6\" s=\"19\"><v>431214</v></c><c r=\"C6\" s=\"20\"><v>41247</v></c><c r=\"D6\" s=\"21\"><v>8567</v></c><c r=\"E6\" s=\"22\"><v>0.05</v></c><c r=\"F6\" s=\"23\"><v>6</v></c><c r=\"G6\" s=\"15\"><v>41247</v></c><c r=\"H6\" s=\"24\"><v>12</v></c><c r=\"I6\" s=\"25\" t=\"s\"><v>12</v></c></row><row r=\"7\" spans=\"1:9\" ht=\"20.5\" customHeight=\"1\"><c r=\"A7\" s=\"18\"><v>5</v></c><c r=\"B7\" s=\"19\"><v>431214</v></c><c r=\"C7\" s=\"20\"><v>41247</v></c><c r=\"D7\" s=\"21\"><v>7123</v></c><c r=\"E7\" s=\"22\"><v>0.05</v></c><c r=\"F7\" s=\"23\"><v>13</v></c><c r=\"G7\" s=\"15\"><v>41247</v></c><c r=\"H7\" s=\"24\"><v>12</v></c><c r=\"I7\" s=\"25\" t=\"s\"><v>12</v></c></row><row r=\"8\" spans=\"1:9\" ht=\"20.25\" customHeight=\"1\"><c r=\"A8\" s=\"18\"><v>6</v></c><c r=\"B8\" s=\"19\"><v>431214</v></c><c r=\"C8\" s=\"20\"><v>41247</v></c><c r=\"D8\" s=\"21\"><v>6000</v></c><c r=\"E8\" s=\"22\"><v>0.05</v></c><c r=\"F8\" s=\"23\"><v>4</v></c><c r=\"G8\" s=\"15\"><v>41247</v></c><c r=\"H8\" s=\"24\"><v>12</v></c><c r=\"I8\" s=\"25\" t=\"s\"><v>12</v></c></row><row r=\"9\" spans=\"1:9\" ht=\"20.25\" customHeight=\"1\"><c r=\"A9\" s=\"18\"><v>7</v></c><c r=\"B9\" s=\"19\"><v>431214</v></c><c r=\"C9\" s=\"20\"><v>41247</v></c><c r=\"D9\" s=\"21\"><v>50000</v></c><c r=\"E9\" s=\"22\"><v>0.05</v></c><c r=\"F9\" s=\"23\"><v>3</v></c><c r=\"G9\" s=\"15\"><v>41247</v></c><c r=\"H9\" s=\"24\"><v>12</v></c><c r=\"I9\" s=\"25\" t=\"s\"><v>12</v></c></row><row r=\"10\" spans=\"1:9\" ht=\"20.25\" customHeight=\"1\"><c r=\"A10\" s=\"18\"><v>8</v></c><c r=\"B10\" s=\"19\"><v>431214</v></c><c r=\"C10\" s=\"20\"><v>41247</v></c><c r=\"D10\" s=\"21\"><v>120000</v></c><c r=\"E10\" s=\"22\"><v>0.05</v></c><c r=\"F10\" s=\"23\"><v>1</v></c><c r=\"G10\" s=\"15\"><v>41247</v></c><c r=\"H10\" s=\"24\"><v>12</v></c><c r=\"I10\" s=\"25\" t=\"s\"><v>12</v></c></row><row r=\"11\" spans=\"1:9\" ht=\"20.25\" customHeight=\"1\"><c r=\"A11\" s=\"18\"><v>9</v></c><c r=\"B11\" s=\"26\"><v>3695064866</v></c><c r=\"C11\" s=\"27\"><v>41247</v></c><c r=\"D11\" s=\"21\"><v>1234</v></c><c r=\"E11\" s=\"22\"><v>0.05</v></c><c r=\"F11\" s=\"23\"><v>2</v></c><c r=\"G11\" s=\"15\"><v>41247</v></c><c r=\"H11\" s=\"24\"><v>12</v></c><c r=\"I11\" s=\"25\" t=\"s\"><v>13</v></c></row><row r=\"12\" spans=\"1:9\" ht=\"20.25\" customHeight=\"1\"><c r=\"A12\" s=\"18\"><v>10</v></c><c r=\"B12\" s=\"26\"><v>3695064866</v></c><c r=\"C12\" s=\"27\"><v>41247</v></c><c r=\"D12\" s=\"21\"><v>122</v></c><c r=\"E12\" s=\"22\"><v>0.05</v></c><c r=\"F12\" s=\"23\"><v>1</v></c><c r=\"G12\" s=\"15\"><v>41247</v></c><c r=\"H12\" s=\"24\"><v>12</v></c><c r=\"I12\" s=\"25\" t=\"s\"><v>13</v></c></row><row r=\"13\" spans=\"1:9\" ht=\"20.25\" customHeight=\"1\"><c r=\"A13\" s=\"18\"><v>11</v></c><c r=\"B13\" s=\"26\"><v>3695064866</v></c><c r=\"C13\" s=\"27\"><v>41247</v></c><c r=\"D13\" s=\"21\"><v>781115.19</v></c><c r=\"E13\" s=\"22\"><v>0.05</v></c><c r=\"F13\" s=\"23\"><v>1</v></c><c r=\"G13\" s=\"15\"><v>41247</v></c><c r=\"H13\" s=\"24\"><v>12</v></c><c r=\"I13\" s=\"25\" t=\"s\"><v>13</v></c></row><row r=\"14\" spans=\"1:9\" ht=\"20.25\" customHeight=\"1\"><c r=\"A14\" s=\"18\"><v>12</v></c><c r=\"B14\" s=\"26\"><v>3695064866</v></c><c r=\"C14\" s=\"27\"><v>41247</v></c><c r=\"D14\" s=\"21\"><v>781115.19</v></c><c r=\"E14\" s=\"22\"><v>0.05</v></c><c r=\"F14\" s=\"23\"><v>1</v></c><c r=\"G14\" s=\"15\"><v>41247</v></c><c r=\"H14\" s=\"24\"><v>12</v></c><c r=\"I14\" s=\"25\" t=\"s\"><v>13</v></c></row><row r=\"15\" spans=\"1:9\" ht=\"20.25\" customHeight=\"1\"><c r=\"A15\" s=\"18\"><v>13</v></c><c r=\"B15\" s=\"26\"><v>3695064866</v></c><c r=\"C15\" s=\"27\"><v>41247</v></c><c r=\"D15\" s=\"21\"><v>781115.19</v></c><c r=\"E15\" s=\"22\"><v>0.05</v></c><c r=\"F15\" s=\"23\"><v>1</v></c><c r=\"G15\" s=\"15\"><v>41247</v></c><c r=\"H15\" s=\"24\"><v>12</v></c><c r=\"I15\" s=\"25\" t=\"s\"><v>13</v></c></row><row r=\"16\" spans=\"1:9\" ht=\"20.25\" customHeight=\"1\"><c r=\"A16\" s=\"18\"><v>14</v></c><c r=\"B16\" s=\"26\"><v>3695064866</v></c><c r=\"C16\" s=\"27\"><v>41247</v></c><c r=\"D16\" s=\"21\"><v>781115.19</v></c><c r=\"E16\" s=\"22\"><v>0.05</v></c><c r=\"F16\" s=\"23\"><v>1</v></c><c r=\"G16\" s=\"15\"><v>41247</v></c><c r=\"H16\" s=\"24\"><v>12</v></c><c r=\"I16\" s=\"25\" t=\"s\"><v>13</v></c></row><row r=\"17\" spans=\"1:9\" ht=\"20.25\" customHeight=\"1\"><c r=\"A17\" s=\"18\"><v>15</v></c><c r=\"B17\" s=\"26\"><v>3695064866</v></c><c r=\"C17\" s=\"27\"><v>41247</v></c><c r=\"D17\" s=\"21\"><v>781115.19</v></c><c r=\"E17\" s=\"22\"><v>0.05</v></c><c r=\"F17\" s=\"23\"><v>1</v></c><c r=\"G17\" s=\"15\"><v>41247</v></c><c r=\"H17\" s=\"24\"><v>12</v></c><c r=\"I17\" s=\"25\" t=\"s\"><v>13</v></c></row><row r=\"18\" spans=\"1:9\" ht=\"20.25\" customHeight=\"1\"><c r=\"A18\" s=\"18\"><v>16</v></c><c r=\"B18\" s=\"26\"><v>3695064866</v></c><c r=\"C18\" s=\"28\"><v>41247</v></c><c r=\"D18\" s=\"29\"><v>781115.19</v></c><c r=\"E18\" s=\"22\"><v>0.05</v></c><c r=\"F18\" s=\"23\"><v>1</v></c><c r=\"G18\" s=\"30\"><v>41247</v></c><c r=\"H18\" s=\"24\"><v>12</v></c><c r=\"I18\" s=\"31\" t=\"s\"><v>13</v></c></row><row r=\"19\" spans=\"1:9\" ht=\"20.25\" customHeight=\"1\"><c r=\"A19\" s=\"18\"/><c r=\"B19\" s=\"32\"/><c r=\"C19\" s=\"33\"/><c r=\"D19\" s=\"33\"/><c r=\"E19\" s=\"33\"/><c r=\"F19\" s=\"33\"/><c r=\"G19\" s=\"33\"/><c r=\"H19\" s=\"33\"/><c r=\"I19\" s=\"33\"/></row><row r=\"20\" spans=\"1:9\" ht=\"20.25\" customHeight=\"1\"><c r=\"A20\" s=\"18\"/><c r=\"B20\" s=\"32\"/><c r=\"C20\" s=\"33\"/><c r=\"D20\" s=\"33\"/><c r=\"E20\" s=\"33\"/><c r=\"F20\" s=\"33\"/><c r=\"G20\" s=\"33\"/><c r=\"H20\" s=\"33\"/><c r=\"I20\" s=\"33\"/></row><row r=\"21\" spans=\"1:9\" ht=\"20.25\" customHeight=\"1\"><c r=\"A21\" s=\"18\"/><c r=\"B21\" s=\"32\"/><c r=\"C21\" s=\"33\"/><c r=\"D21\" s=\"33\"/><c r=\"E21\" s=\"33\"/><c r=\"F21\" s=\"33\"/><c r=\"G21\" s=\"33\"/><c r=\"H21\" s=\"33\"/><c r=\"I21\" s=\"33\"/></row><row r=\"22\" spans=\"1:9\" ht=\"20.25\" customHeight=\"1\"><c r=\"A22\" s=\"18\"/><c r=\"B22\" s=\"32\"/><c r=\"C22\" s=\"33\"/><c r=\"D22\" s=\"33\"/><c r=\"E22\" s=\"33\"/><c r=\"F22\" s=\"33\"/><c r=\"G22\" s=\"33\"/><c r=\"H22\" s=\"33\"/><c r=\"I22\" s=\"33\"/></row><row r=\"23\" spans=\"1:9\" ht=\"20.25\" customHeight=\"1\"><c r=\"A23\" s=\"18\"/><c r=\"B23\" s=\"32\"/><c r=\"C23\" s=\"33\"/><c r=\"D23\" s=\"33\"/><c r=\"E23\" s=\"33\"/><c r=\"F23\" s=\"33\"/><c r=\"G23\" s=\"33\"/><c r=\"H23\" s=\"33\"/><c r=\"I23\" s=\"33\"/></row></sheetData><mergeCells count=\"1\"><mergeCell ref=\"A1:I1\"/></mergeCells><phoneticPr fontId=\"5\" type=\"noConversion\"/><pageMargins left=\"0.5\" right=\"0.5\" top=\"0.75\" bottom=\"0.75\" header=\"0.27777779102325439\" footer=\"0.27777779102325439\"/><pageSetup paperSize=\"9\" orientation=\"portrait\" horizontalDpi=\"4294967292\" verticalDpi=\"4294967292\"/><headerFooter><oddFooter>&amp;C&amp;\"Helvetica,Regular\"&amp;12&amp;K000000&amp;P</oddFooter></headerFooter><extLst><ext uri=\"{64002731-A6B0-56B0-2670-7721B7C09600}\" xmlns:mx=\"http://schemas.microsoft.com/office/mac/excel/2008/main\"><mx:PLV Mode=\"0\" OnePage=\"0\" WScale=\"0\"/></ext></extLst></worksheet>";
		
		/*
		 * <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
		 * <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">
		 * <dimension ref="A1:I17"/>
		 * <sheetViews>
		 *   <sheetView tabSelected="1" workbookViewId="0">
		 *     <selection activeCell="D21" sqref="D21"/>
		 *   </sheetView>
		 * </sheetViews>
		 * <sheetFormatPr baseColWidth="10" defaultRowHeight="12" x14ac:dyDescent="0"/>
		 * <sheetData>
		 *   <row r="1" spans="1:9" s="1" customFormat="1" ht="23" customHeight="1">
		 *     <c r="A1" s="2" t="s"><v>1</v></c>
		 *     <c r="B1" s="3" t="s"><v>2</v></c>
		 *     <c r="C1" s="4" t="s"><v>3</v></c>
		 *     <c r="D1" s="5" t="s"><v>4</v></c>
		 *     <c r="E1" s="6" t="s"><v>5</v></c>
		 *     <c r="F1" s="7" t="s"><v>6</v></c>
		 *     <c r="G1" s="8" t="s"><v>7</v></c>
		 *     <c r="H1" s="2" t="s"><v>8</v></c>
		 *     <c r="I1" s="4" t="s"><v>9</v></c>
		 *   </row>
		 *   <row r="2" spans="1:9" s="1" customFormat="1" ht="21" customHeight="1">
		 *     <c r="A2" s="9"><v>1</v></c>
		 *     <c r="B2" s="10"><v>431212</v></c>
		 *     <c r="C2" s="11"><v>41242</v></c>
		 *     <c r="D2" s="12"><v>11</v></c>
		 *     <c r="E2" s="13"><v>0.05</v></c>
		 *     <c r="F2" s="14"><v>1</v></c>
		 *     <c r="G2" s="15"><v>41242</v></c>
		 *     <c r="H2" s="16"><v>12</v></c>
		 *     <c r="I2" s="17" t="s"><v>10</v></c>
		 *   </row><row r="3" spans="1:9" s="1" customFormat="1" ht="21" customHeight="1"><c r="A3" s="18"><v>2</v></c><c r="B3" s="19"><v>6666</v></c><c r="C3" s="20"><v>41243</v></c><c r="D3" s="21"><v>753578.16599999997</v></c><c r="E3" s="22"><v>0.05</v></c><c r="F3" s="23"><v>1</v></c><c r="G3" s="15"><v>41243</v></c><c r="H3" s="24"><v>12</v></c><c r="I3" s="25" t="s"><v>11</v></c></row><row r="4" spans="1:9" s="1" customFormat="1" ht="21" customHeight="1"><c r="A4" s="18"><v>3</v></c><c r="B4" s="19"><v>431214</v></c><c r="C4" s="20"><v>41244</v></c><c r="D4" s="21"><v>3452</v></c><c r="E4" s="22"><v>0.05</v></c><c r="F4" s="23"><v>2</v></c><c r="G4" s="15"><v>41244</v></c><c r="H4" s="24"><v>12</v></c><c r="I4" s="25" t="s"><v>12</v></c></row><row r="5" spans="1:9" s="1" customFormat="1" ht="21" customHeight="1"><c r="A5" s="18"><v>4</v></c><c r="B5" s="19"><v>431214</v></c><c r="C5" s="20"><v>41247</v></c><c r="D5" s="21"><v>8567</v></c><c r="E5" s="22"><v>0.05</v></c><c r="F5" s="23"><v>6</v></c><c r="G5" s="15"><v>41247</v></c><c r="H5" s="24"><v>12</v></c><c r="I5" s="25" t="s"><v>12</v></c></row><row r="6" spans="1:9" s="1" customFormat="1" ht="20.5" customHeight="1"><c r="A6" s="18"><v>5</v></c><c r="B6" s="19"><v>431214</v></c><c r="C6" s="20"><v>41247</v></c><c r="D6" s="21"><v>7123</v></c><c r="E6" s="22"><v>0.05</v></c><c r="F6" s="23"><v>13</v></c><c r="G6" s="15"><v>41247</v></c><c r="H6" s="24"><v>12</v></c><c r="I6" s="25" t="s"><v>12</v></c></row><row r="7" spans="1:9" s="1" customFormat="1" ht="20.25" customHeight="1"><c r="A7" s="18"><v>6</v></c><c r="B7" s="19"><v>431214</v></c><c r="C7" s="20"><v>41247</v></c><c r="D7" s="21"><v>6000</v></c><c r="E7" s="22"><v>0.05</v></c><c r="F7" s="23"><v>4</v></c><c r="G7" s="15"><v>41247</v></c><c r="H7" s="24"><v>12</v></c><c r="I7" s="25" t="s"><v>12</v></c></row><row r="8" spans="1:9" s="1" customFormat="1" ht="20.25" customHeight="1"><c r="A8" s="18"><v>7</v></c><c r="B8" s="19"><v>431214</v></c><c r="C8" s="20"><v>41247</v></c><c r="D8" s="21"><v>50000</v></c><c r="E8" s="22"><v>0.05</v></c><c r="F8" s="23"><v>3</v></c><c r="G8" s="15"><v>41247</v></c><c r="H8" s="24"><v>12</v></c><c r="I8" s="25" t="s"><v>12</v></c></row><row r="9" spans="1:9" s="1" customFormat="1" ht="20.25" customHeight="1"><c r="A9" s="18"><v>8</v></c><c r="B9" s="19"><v>431214</v></c><c r="C9" s="20"><v>41247</v></c><c r="D9" s="21"><v>120000</v></c><c r="E9" s="22"><v>0.05</v></c><c r="F9" s="23"><v>1</v></c><c r="G9" s="15"><v>41247</v></c><c r="H9" s="24"><v>12</v></c><c r="I9" s="25" t="s"><v>12</v></c></row><row r="10" spans="1:9" s="1" customFormat="1" ht="20.25" customHeight="1"><c r="A10" s="18"><v>9</v></c><c r="B10" s="26"><v>3695064866</v></c><c r="C10" s="27"><v>41247</v></c><c r="D10" s="21"><v>1234</v></c><c r="E10" s="22"><v>0.05</v></c><c r="F10" s="23"><v>2</v></c><c r="G10" s="15"><v>41247</v></c><c r="H10" s="24"><v>12</v></c><c r="I10" s="25" t="s"><v>13</v></c></row><row r="11" spans="1:9" s="1" customFormat="1" ht="20.25" customHeight="1"><c r="A11" s="18"><v>10</v></c><c r="B11" s="26"><v>3695064866</v></c><c r="C11" s="27"><v>41247</v></c><c r="D11" s="21"><v>122</v></c><c r="E11" s="22"><v>0.05</v></c><c r="F11" s="23"><v>1</v></c><c r="G11" s="15"><v>41247</v></c><c r="H11" s="24"><v>12</v></c><c r="I11" s="25" t="s"><v>13</v></c></row><row r="12" spans="1:9" s="1" customFormat="1" ht="20.25" customHeight="1"><c r="A12" s="18"><v>11</v></c><c r="B12" s="26"><v>3695064866</v></c><c r="C12" s="27"><v>41247</v></c><c r="D12" s="21"><v>781115.19</v></c><c r="E12" s="22"><v>0.05</v></c><c r="F12" s="23"><v>1</v></c><c r="G12" s="15"><v>41247</v></c><c r="H12" s="24"><v>12</v></c><c r="I12" s="25" t="s"><v>13</v></c></row><row r="13" spans="1:9" s="1" customFormat="1" ht="20.25" customHeight="1"><c r="A13" s="18"><v>12</v></c><c r="B13" s="26"><v>3695064866</v></c><c r="C13" s="27"><v>41247</v></c><c r="D13" s="21"><v>781115.19</v></c><c r="E13" s="22"><v>0.05</v></c><c r="F13" s="23"><v>1</v></c><c r="G13" s="15"><v>41247</v></c><c r="H13" s="24"><v>12</v></c><c r="I13" s="25" t="s"><v>13</v></c></row><row r="14" spans="1:9" s="1" customFormat="1" ht="20.25" customHeight="1"><c r="A14" s="18"><v>13</v></c><c r="B14" s="26"><v>3695064866</v></c><c r="C14" s="27"><v>41247</v></c><c r="D14" s="21"><v>781115.19</v></c><c r="E14" s="22"><v>0.05</v></c><c r="F14" s="23"><v>1</v></c><c r="G14" s="15"><v>41247</v></c><c r="H14" s="24"><v>12</v></c><c r="I14" s="25" t="s"><v>13</v></c></row><row r="15" spans="1:9" s="1" customFormat="1" ht="20.25" customHeight="1"><c r="A15" s="18"><v>14</v></c><c r="B15" s="26"><v>3695064866</v></c><c r="C15" s="27"><v>41247</v></c><c r="D15" s="21"><v>781115.19</v></c><c r="E15" s="22"><v>0.05</v></c><c r="F15" s="23"><v>1</v></c><c r="G15" s="15"><v>41247</v></c><c r="H15" s="24"><v>12</v></c><c r="I15" s="25" t="s"><v>13</v></c></row><row r="16" spans="1:9" s="1" customFormat="1" ht="20.25" customHeight="1"><c r="A16" s="18"><v>15</v></c><c r="B16" s="26"><v>3695064866</v></c><c r="C16" s="27"><v>41247</v></c><c r="D16" s="21"><v>781115.19</v></c><c r="E16" s="22"><v>0.05</v></c><c r="F16" s="23"><v>1</v></c><c r="G16" s="15"><v>41247</v></c><c r="H16" s="24"><v>12</v></c><c r="I16" s="25" t="s"><v>13</v></c></row><row r="17" spans="1:9" s="1" customFormat="1" ht="20.25" customHeight="1"><c r="A17" s="18"><v>16</v></c><c r="B17" s="26"><v>3695064866</v></c><c r="C17" s="28"><v>41247</v></c><c r="D17" s="29"><v>781115.19</v></c><c r="E17" s="22"><v>0.05</v></c><c r="F17" s="23"><v>1</v></c><c r="G17" s="30"><v>41247</v></c><c r="H17" s="24"><v>12</v></c><c r="I17" s="31" t="s"><v>13</v></c></row></sheetData><phoneticPr fontId="5" type="noConversion"/><pageMargins left="0.75" right="0.75" top="1" bottom="1" header="0.5" footer="0.5"/><extLst><ext uri="{64002731-A6B0-56B0-2670-7721B7C09600}" xmlns:mx="http://schemas.microsoft.com/office/mac/excel/2008/main"><mx:PLV Mode="0" OnePage="0" WScale="0"/></ext></extLst></worksheet>
		 */
		sheet2Xml = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n" + 
				"<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" mc:Ignorable=\"x14ac\" xmlns:x14ac=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac\"><dimension ref=\"A1:I17\"/><sheetViews><sheetView tabSelected=\"1\" workbookViewId=\"0\"><selection activeCell=\"D21\" sqref=\"D21\"/></sheetView></sheetViews><sheetFormatPr baseColWidth=\"10\" defaultRowHeight=\"12\" x14ac:dyDescent=\"0\"/><sheetData><row r=\"1\" spans=\"1:9\" s=\"1\" customFormat=\"1\" ht=\"23\" customHeight=\"1\"><c r=\"A1\" s=\"2\" t=\"s\"><v>1</v></c><c r=\"B1\" s=\"3\" t=\"s\"><v>2</v></c><c r=\"C1\" s=\"4\" t=\"s\"><v>3</v></c><c r=\"D1\" s=\"5\" t=\"s\"><v>4</v></c><c r=\"E1\" s=\"6\" t=\"s\"><v>5</v></c><c r=\"F1\" s=\"7\" t=\"s\"><v>6</v></c><c r=\"G1\" s=\"8\" t=\"s\"><v>7</v></c><c r=\"H1\" s=\"2\" t=\"s\"><v>8</v></c><c r=\"I1\" s=\"4\" t=\"s\"><v>9</v></c></row><row r=\"2\" spans=\"1:9\" s=\"1\" customFormat=\"1\" ht=\"21\" customHeight=\"1\"><c r=\"A2\" s=\"9\"><v>1</v></c><c r=\"B2\" s=\"10\"><v>431212</v></c><c r=\"C2\" s=\"11\"><v>41242</v></c><c r=\"D2\" s=\"12\"><v>11</v></c><c r=\"E2\" s=\"13\"><v>0.05</v></c><c r=\"F2\" s=\"14\"><v>1</v></c><c r=\"G2\" s=\"15\"><v>41242</v></c><c r=\"H2\" s=\"16\"><v>12</v></c><c r=\"I2\" s=\"17\" t=\"s\"><v>10</v></c></row><row r=\"3\" spans=\"1:9\" s=\"1\" customFormat=\"1\" ht=\"21\" customHeight=\"1\"><c r=\"A3\" s=\"18\"><v>2</v></c><c r=\"B3\" s=\"19\"><v>6666</v></c><c r=\"C3\" s=\"20\"><v>41243</v></c><c r=\"D3\" s=\"21\"><v>753578.16599999997</v></c><c r=\"E3\" s=\"22\"><v>0.05</v></c><c r=\"F3\" s=\"23\"><v>1</v></c><c r=\"G3\" s=\"15\"><v>41243</v></c><c r=\"H3\" s=\"24\"><v>12</v></c><c r=\"I3\" s=\"25\" t=\"s\"><v>11</v></c></row><row r=\"4\" spans=\"1:9\" s=\"1\" customFormat=\"1\" ht=\"21\" customHeight=\"1\"><c r=\"A4\" s=\"18\"><v>3</v></c><c r=\"B4\" s=\"19\"><v>431214</v></c><c r=\"C4\" s=\"20\"><v>41244</v></c><c r=\"D4\" s=\"21\"><v>3452</v></c><c r=\"E4\" s=\"22\"><v>0.05</v></c><c r=\"F4\" s=\"23\"><v>2</v></c><c r=\"G4\" s=\"15\"><v>41244</v></c><c r=\"H4\" s=\"24\"><v>12</v></c><c r=\"I4\" s=\"25\" t=\"s\"><v>12</v></c></row><row r=\"5\" spans=\"1:9\" s=\"1\" customFormat=\"1\" ht=\"21\" customHeight=\"1\"><c r=\"A5\" s=\"18\"><v>4</v></c><c r=\"B5\" s=\"19\"><v>431214</v></c><c r=\"C5\" s=\"20\"><v>41247</v></c><c r=\"D5\" s=\"21\"><v>8567</v></c><c r=\"E5\" s=\"22\"><v>0.05</v></c><c r=\"F5\" s=\"23\"><v>6</v></c><c r=\"G5\" s=\"15\"><v>41247</v></c><c r=\"H5\" s=\"24\"><v>12</v></c><c r=\"I5\" s=\"25\" t=\"s\"><v>12</v></c></row><row r=\"6\" spans=\"1:9\" s=\"1\" customFormat=\"1\" ht=\"20.5\" customHeight=\"1\"><c r=\"A6\" s=\"18\"><v>5</v></c><c r=\"B6\" s=\"19\"><v>431214</v></c><c r=\"C6\" s=\"20\"><v>41247</v></c><c r=\"D6\" s=\"21\"><v>7123</v></c><c r=\"E6\" s=\"22\"><v>0.05</v></c><c r=\"F6\" s=\"23\"><v>13</v></c><c r=\"G6\" s=\"15\"><v>41247</v></c><c r=\"H6\" s=\"24\"><v>12</v></c><c r=\"I6\" s=\"25\" t=\"s\"><v>12</v></c></row><row r=\"7\" spans=\"1:9\" s=\"1\" customFormat=\"1\" ht=\"20.25\" customHeight=\"1\"><c r=\"A7\" s=\"18\"><v>6</v></c><c r=\"B7\" s=\"19\"><v>431214</v></c><c r=\"C7\" s=\"20\"><v>41247</v></c><c r=\"D7\" s=\"21\"><v>6000</v></c><c r=\"E7\" s=\"22\"><v>0.05</v></c><c r=\"F7\" s=\"23\"><v>4</v></c><c r=\"G7\" s=\"15\"><v>41247</v></c><c r=\"H7\" s=\"24\"><v>12</v></c><c r=\"I7\" s=\"25\" t=\"s\"><v>12</v></c></row><row r=\"8\" spans=\"1:9\" s=\"1\" customFormat=\"1\" ht=\"20.25\" customHeight=\"1\"><c r=\"A8\" s=\"18\"><v>7</v></c><c r=\"B8\" s=\"19\"><v>431214</v></c><c r=\"C8\" s=\"20\"><v>41247</v></c><c r=\"D8\" s=\"21\"><v>50000</v></c><c r=\"E8\" s=\"22\"><v>0.05</v></c><c r=\"F8\" s=\"23\"><v>3</v></c><c r=\"G8\" s=\"15\"><v>41247</v></c><c r=\"H8\" s=\"24\"><v>12</v></c><c r=\"I8\" s=\"25\" t=\"s\"><v>12</v></c></row><row r=\"9\" spans=\"1:9\" s=\"1\" customFormat=\"1\" ht=\"20.25\" customHeight=\"1\"><c r=\"A9\" s=\"18\"><v>8</v></c><c r=\"B9\" s=\"19\"><v>431214</v></c><c r=\"C9\" s=\"20\"><v>41247</v></c><c r=\"D9\" s=\"21\"><v>120000</v></c><c r=\"E9\" s=\"22\"><v>0.05</v></c><c r=\"F9\" s=\"23\"><v>1</v></c><c r=\"G9\" s=\"15\"><v>41247</v></c><c r=\"H9\" s=\"24\"><v>12</v></c><c r=\"I9\" s=\"25\" t=\"s\"><v>12</v></c></row><row r=\"10\" spans=\"1:9\" s=\"1\" customFormat=\"1\" ht=\"20.25\" customHeight=\"1\"><c r=\"A10\" s=\"18\"><v>9</v></c><c r=\"B10\" s=\"26\"><v>3695064866</v></c><c r=\"C10\" s=\"27\"><v>41247</v></c><c r=\"D10\" s=\"21\"><v>1234</v></c><c r=\"E10\" s=\"22\"><v>0.05</v></c><c r=\"F10\" s=\"23\"><v>2</v></c><c r=\"G10\" s=\"15\"><v>41247</v></c><c r=\"H10\" s=\"24\"><v>12</v></c><c r=\"I10\" s=\"25\" t=\"s\"><v>13</v></c></row><row r=\"11\" spans=\"1:9\" s=\"1\" customFormat=\"1\" ht=\"20.25\" customHeight=\"1\"><c r=\"A11\" s=\"18\"><v>10</v></c><c r=\"B11\" s=\"26\"><v>3695064866</v></c><c r=\"C11\" s=\"27\"><v>41247</v></c><c r=\"D11\" s=\"21\"><v>122</v></c><c r=\"E11\" s=\"22\"><v>0.05</v></c><c r=\"F11\" s=\"23\"><v>1</v></c><c r=\"G11\" s=\"15\"><v>41247</v></c><c r=\"H11\" s=\"24\"><v>12</v></c><c r=\"I11\" s=\"25\" t=\"s\"><v>13</v></c></row><row r=\"12\" spans=\"1:9\" s=\"1\" customFormat=\"1\" ht=\"20.25\" customHeight=\"1\"><c r=\"A12\" s=\"18\"><v>11</v></c><c r=\"B12\" s=\"26\"><v>3695064866</v></c><c r=\"C12\" s=\"27\"><v>41247</v></c><c r=\"D12\" s=\"21\"><v>781115.19</v></c><c r=\"E12\" s=\"22\"><v>0.05</v></c><c r=\"F12\" s=\"23\"><v>1</v></c><c r=\"G12\" s=\"15\"><v>41247</v></c><c r=\"H12\" s=\"24\"><v>12</v></c><c r=\"I12\" s=\"25\" t=\"s\"><v>13</v></c></row><row r=\"13\" spans=\"1:9\" s=\"1\" customFormat=\"1\" ht=\"20.25\" customHeight=\"1\"><c r=\"A13\" s=\"18\"><v>12</v></c><c r=\"B13\" s=\"26\"><v>3695064866</v></c><c r=\"C13\" s=\"27\"><v>41247</v></c><c r=\"D13\" s=\"21\"><v>781115.19</v></c><c r=\"E13\" s=\"22\"><v>0.05</v></c><c r=\"F13\" s=\"23\"><v>1</v></c><c r=\"G13\" s=\"15\"><v>41247</v></c><c r=\"H13\" s=\"24\"><v>12</v></c><c r=\"I13\" s=\"25\" t=\"s\"><v>13</v></c></row><row r=\"14\" spans=\"1:9\" s=\"1\" customFormat=\"1\" ht=\"20.25\" customHeight=\"1\"><c r=\"A14\" s=\"18\"><v>13</v></c><c r=\"B14\" s=\"26\"><v>3695064866</v></c><c r=\"C14\" s=\"27\"><v>41247</v></c><c r=\"D14\" s=\"21\"><v>781115.19</v></c><c r=\"E14\" s=\"22\"><v>0.05</v></c><c r=\"F14\" s=\"23\"><v>1</v></c><c r=\"G14\" s=\"15\"><v>41247</v></c><c r=\"H14\" s=\"24\"><v>12</v></c><c r=\"I14\" s=\"25\" t=\"s\"><v>13</v></c></row><row r=\"15\" spans=\"1:9\" s=\"1\" customFormat=\"1\" ht=\"20.25\" customHeight=\"1\"><c r=\"A15\" s=\"18\"><v>14</v></c><c r=\"B15\" s=\"26\"><v>3695064866</v></c><c r=\"C15\" s=\"27\"><v>41247</v></c><c r=\"D15\" s=\"21\"><v>781115.19</v></c><c r=\"E15\" s=\"22\"><v>0.05</v></c><c r=\"F15\" s=\"23\"><v>1</v></c><c r=\"G15\" s=\"15\"><v>41247</v></c><c r=\"H15\" s=\"24\"><v>12</v></c><c r=\"I15\" s=\"25\" t=\"s\"><v>13</v></c></row><row r=\"16\" spans=\"1:9\" s=\"1\" customFormat=\"1\" ht=\"20.25\" customHeight=\"1\"><c r=\"A16\" s=\"18\"><v>15</v></c><c r=\"B16\" s=\"26\"><v>3695064866</v></c><c r=\"C16\" s=\"27\"><v>41247</v></c><c r=\"D16\" s=\"21\"><v>781115.19</v></c><c r=\"E16\" s=\"22\"><v>0.05</v></c><c r=\"F16\" s=\"23\"><v>1</v></c><c r=\"G16\" s=\"15\"><v>41247</v></c><c r=\"H16\" s=\"24\"><v>12</v></c><c r=\"I16\" s=\"25\" t=\"s\"><v>13</v></c></row><row r=\"17\" spans=\"1:9\" s=\"1\" customFormat=\"1\" ht=\"20.25\" customHeight=\"1\"><c r=\"A17\" s=\"18\"><v>16</v></c><c r=\"B17\" s=\"26\"><v>3695064866</v></c><c r=\"C17\" s=\"28\"><v>41247</v></c><c r=\"D17\" s=\"29\"><v>781115.19</v></c><c r=\"E17\" s=\"22\"><v>0.05</v></c><c r=\"F17\" s=\"23\"><v>1</v></c><c r=\"G17\" s=\"30\"><v>41247</v></c><c r=\"H17\" s=\"24\"><v>12</v></c><c r=\"I17\" s=\"31\" t=\"s\"><v>13</v></c></row></sheetData><phoneticPr fontId=\"5\" type=\"noConversion\"/><pageMargins left=\"0.75\" right=\"0.75\" top=\"1\" bottom=\"1\" header=\"0.5\" footer=\"0.5\"/><extLst><ext uri=\"{64002731-A6B0-56B0-2670-7721B7C09600}\" xmlns:mx=\"http://schemas.microsoft.com/office/mac/excel/2008/main\"><mx:PLV Mode=\"0\" OnePage=\"0\" WScale=\"0\"/></ext></extLst></worksheet>";
	}

	@Test
	public void testSharedStringsXML() throws SAXException, ParserConfigurationException, IOException {
		InputStream inputStream = new ByteArrayInputStream(sharedStringsXml.getBytes());
		InputSource inputSource = new InputSource(inputStream);
		
		xmlReader.setContentHandler(new OneRowHandler());
		xmlReader.parse(inputSource);
		
		inputStream.close();
	}
	
	@Test
	public void testWorkbookXML() throws SAXException, ParserConfigurationException, IOException {
		InputStream inputStream = new ByteArrayInputStream(workbookXml.getBytes());
		InputSource inputSource = new InputSource(inputStream);
		
		xmlReader.setContentHandler(new OneRowHandler());
		xmlReader.parse(inputSource);
		
		inputStream.close();
	}
	
	@Test
	public void testSheet1XML() throws SAXException, ParserConfigurationException, IOException {
		InputStream inputStream = new ByteArrayInputStream(sheet1Xml.getBytes());
		InputSource inputSource = new InputSource(inputStream);
		
		xmlReader.setContentHandler(new OneRowHandler());
		xmlReader.parse(inputSource);
		
		inputStream.close();
	}
	
	@Test
	public void testSheet2XML() throws SAXException, ParserConfigurationException, IOException {
		InputStream inputStream = new ByteArrayInputStream(sheet2Xml.getBytes());
		InputSource inputSource = new InputSource(inputStream);
		
		xmlReader.setContentHandler(new OneRowHandler());
		xmlReader.parse(inputSource);
		
		inputStream.close();
	}
	
	class OneRowHandler extends DefaultHandler {
		@Override
		public InputSource resolveEntity (String publicId, String systemId) throws IOException, SAXException {
			System.out.println("***{ resolveEntity }***");
			System.out.println(" publicId: " + publicId + ", systemId: " + systemId + ".");
			return null;
		}
		
		@Override
		public void notationDecl (String name, String publicId, String systemId) throws SAXException {
			System.out.println("***{ notationDecl }***");
			System.out.println(" name: " + name + ", publicId: " + publicId + ", systemId: " + systemId + ".");
		}
		
		@Override
		public void unparsedEntityDecl (String name, String publicId, String systemId, String notationName) throws SAXException {
			System.out.println("***{ unparsedEntityDecl }***");
			System.out.println(" name: " + name + ", publicId: " + publicId + ", systemId: " + systemId + ", notationName: " + notationName + ".");
		}
		
		@Override
		public void setDocumentLocator (Locator locator) {
			System.out.println("***{ setDocumentLocator }***");
			System.out.println(" locator: " + locator + ".");
		}
		
		@Override
		public void startDocument () throws SAXException {
			System.out.println("***{ startDocument }***");
			System.out.println(" ==>>");
		}
		
		@Override
		public void endDocument () throws SAXException {
			System.out.println("***{ endDocument }***");
			System.out.println(" <<==");
		}
		
		@Override
		public void startPrefixMapping (String prefix, String uri) throws SAXException {
			System.out.println("***{ startPrefixMapping }***");
			System.out.println(" prefix: " + prefix + ", uri: " + uri + ".");
		}
		
		@Override
		public void endPrefixMapping (String prefix) throws SAXException {
			System.out.println("***{ endPrefixMapping }***");
			System.out.println(" prefix: " + prefix);
		}
		
		@Override
		public void startElement (String uri, String localName, String qName, Attributes attributes) throws SAXException {
			System.out.println("***{ startElement }***");
			System.out.println(" uri: " + uri + ", localName: " + localName + ", qName: " + qName + ", Attributes: " + attributes + ".");
		}
		
		@Override
		public void endElement (String uri, String localName, String qName) throws SAXException {
			System.out.println("***{ endElement }***");
			System.out.println(" uri: " + uri + ", localName: " + localName + ", qName: " + qName + ".");
		}

		@Override
		public void characters (char ch[], int start, int length) throws SAXException {
			System.out.println("***{ characters }***");
			System.out.println("  ch[]: " + ch + ", start: " + start + ", length: " + length + ".");
			System.out.println("  --> [ " + new String(ch, start, length) + " ].");
		}
		
		@Override
		public void ignorableWhitespace (char ch[], int start, int length) throws SAXException {
			System.out.println("***{ ignorableWhitespace }***");
			System.out.println("  ch[]: " + ch + ", start: " + start + ", length: " + length + ".");
			System.out.println("  --> whitespace[ " + new String(ch, start, length) + " ].");
		}
		
		@Override
		public void processingInstruction (String target, String data) throws SAXException {
			System.out.println("***{ processingInstruction }***");
			System.out.println("target: " + target + ", data: " + data + ".");
		}
		
		@Override
		public void skippedEntity (String name) throws SAXException {
			System.out.println("***{ skippedEntity }***");
			System.out.println("name: " + name + ".");
		}
		
		@Override
		public void warning (SAXParseException e) throws SAXException {
			System.out.println("***{ warning }***");
			System.out.println("SAXParseException: " + e + ".");
		}
		
		@Override
		public void error (SAXParseException e) throws SAXException {
			System.out.println("***{ error }***");
			System.out.println("SAXParseException: " + e + ".");
		}
		
		public void fatalError (SAXParseException e) throws SAXException {
			System.out.println("***{ fatalError }***");
			System.out.println("SAXParseException: " + e + ".");
		}
	}
	
	@Test
	public void testExcelSSUserModel() throws EncryptedDocumentException, InvalidFormatException, FileNotFoundException, IOException {
		Workbook wb = WorkbookFactory.create(new FileInputStream("/Users/zoubo/Work/BOC/WorkReport/Weekly report-20190714-邹波-核心.xlsx"));

		DataFormatter formatter = new DataFormatter();
		Sheet sheet = wb.getSheet("周报统计");
		for (Row row : sheet) {
			for (Cell cell : row) {
				CellReference cellRef = new CellReference(row.getRowNum(), cell.getColumnIndex());
	            System.out.print(cellRef.formatAsString());
	            System.out.print(" - ");

	            // get the text that appears in the cell by getting the cell value and applying any data formats (Date, 0.00, 1.23e9, $1.23, etc)
	            String text = formatter.formatCellValue(cell);
	            System.out.println(text);

	            // Alternatively, get the value and format it yourself
	            switch (cell.getCellTypeEnum()) {
	                case STRING:
	                    System.out.println(cell.getRichStringCellValue().getString());
	                    break;
	                case NUMERIC:
	                    if (DateUtil.isCellDateFormatted(cell)) {
	                        System.out.println(cell.getDateCellValue());
	                    } else {
	                        System.out.println(cell.getNumericCellValue());
	                    }
	                    break;
	                case BOOLEAN:
	                    System.out.println(cell.getBooleanCellValue());
	                    break;
	                case FORMULA:
	                    System.out.println(cell.getCellFormula());
	                    break;
	                case BLANK:
	                    System.out.println();
	                    break;
	                default:
	                    System.out.println();
	            }
			}
		}
	}
	
}
