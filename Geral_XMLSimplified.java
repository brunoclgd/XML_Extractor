import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;

import org.apache.commons.collections4.bag.SynchronizedSortedBag;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.w3c.dom.Document;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;
import org.xml.sax.SAXException;

public class Geral_XMLSimplified {
	
	private static int indice = 0;
	private static int line = 1;
	private static int qtdProd = 0;
	private static int inicialPos = 0;
	private static boolean hasCompl = false;
	private static String hasST = null;
	
	public static void copyLineAbove(Sheet sheet, Row row, int actualIndex) {
		for(int i = 0; i < inicialPos; i++) {
			row.createCell(i).setCellValue(sheet.getRow(line-1).getCell(i).toString());
			
				
		}
	}
	
	public static boolean match(String[] fieldsToSave, String key, Sheet sheet, String parentNode) {
		String title;
		for(int i = 0; i < fieldsToSave.length; i++) {
			if(key.equalsIgnoreCase(fieldsToSave[i])) {
				if(key.equals("cProd")) {
					System.out.println("achou o produto");
					qtdProd++;
				}
				if(hasST == null && parentNode.equals("det") && (key.equals("vBCST") || key.equals("vBCSTRet"))) {
					hasST = key;
				}
				else if(hasST != null) {
					title = sheet.getRow(0).getCell(indice).getStringCellValue();
					if((title.equals("vBCST") && key.equals("vBCSTRet")) ||
							(title.equals("vBCSTRet") && key.equals("vBCSST"))) {
						indice++;
					}
				}
				return true;
			}
		}
		return false;
	}
	public static boolean isEmpty(Row row) {
		
		if(row == null ) {
			return true;
		}
		else {
			return false;	
		}
		
	}
	
	private static String searchForBCComplementary(NodeList node, String key) {
		String value = null;
		String text;
		char character;
		int inicialPos, pos, finalPos = 0;
		text = node.item(0).getFirstChild().getTextContent();
		pos = text.indexOf(key);
		if(pos != -1) {
			character = text.charAt(pos);
			inicialPos = pos+key.length()+4;
			finalPos = inicialPos;
			while(character != 'I') {
				character = text.charAt(finalPos);
				finalPos = finalPos + 1;
			}
			value = text.substring(inicialPos, finalPos-2);
		}
		
		return value;
		
	}
	
	private static void ConvertXMLToXLSX(String pathName, OutputStream fileOut, Sheet sheet, Row rowTitle, ArrayList<String> titles) throws ParserConfigurationException, SAXException, IOException {
		File inputFile = new File(pathName);
		DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
		DocumentBuilder builder = factory.newDocumentBuilder();
		Document doc = builder.parse(inputFile);
		doc.getDocumentElement().normalize();
		String[] nodesToLookTo = {"ide", "total", "det"};
		String[] fieldsToSave = {"natOp","nNf", "dEmi", "cProd","xProd","qCom", "vUnCom", "vProd", "vBCST", "vBCSTRet", "vNF"};
		Row rowData = sheet.createRow(line);
		for(int i = 0; i < nodesToLookTo.length; i++) {
			NodeList firstNode = doc.getElementsByTagName(nodesToLookTo[i]);
			System.out.println("------Label "+nodesToLookTo[i]);
			
			for(int x = 0; x < firstNode.getLength(); x++) {
				
				if(nodesToLookTo[i].equalsIgnoreCase("det") && qtdProd > 0) {
					indice = inicialPos;
					line++;
					rowData = sheet.createRow(line);
					copyLineAbove(sheet, rowData, indice);
				}
				else if(qtdProd == 0) {
					inicialPos = indice;
				}
				
				System.out.println("***Produto "+(x+1)+" ***");
				Node firstChildNode = firstNode.item(x).getFirstChild();
				if(firstChildNode.getFirstChild() == null) {
					
					while(firstChildNode != null) {
	
						if(firstChildNode.getFirstChild() != null && firstChildNode.hasChildNodes()) {
							
							Node secondChildNode = firstChildNode.getFirstChild();
							while(secondChildNode != null) {
								if(match(fieldsToSave, secondChildNode.getNodeName(), sheet, nodesToLookTo[i])) {
									rowTitle.createCell(indice).setCellValue(secondChildNode.getNodeName());
									rowData.createCell(indice).setCellValue(secondChildNode.getTextContent());
									indice++;
									System.out.println(secondChildNode.getNodeName()+" = "+secondChildNode.getTextContent());
								}
								
								secondChildNode = secondChildNode.getNextSibling();
							}
						}
						else {
							if(match(fieldsToSave, firstChildNode.getNodeName(), sheet, nodesToLookTo[i])) {
								rowTitle.createCell(indice).setCellValue(firstChildNode.getNodeName());
								rowData.createCell(indice).setCellValue(firstChildNode.getTextContent());
								indice++;
								System.out.println(firstChildNode.getNodeName()+" = "+firstChildNode.getTextContent());	
							}
							
						}
						firstChildNode = firstChildNode.getNextSibling();
						
					}
					
				}
				else if(firstChildNode.hasChildNodes()) {
					
					while (firstChildNode != null) {
						
						if(firstChildNode.getFirstChild() == null) {
							
							while(firstChildNode != null && firstChildNode.getFirstChild() == null) {
								
								if(firstChildNode.getFirstChild() != null) {
									
									Node secondaryChildNode = firstChildNode.getFirstChild();
									
									while(secondaryChildNode != null) {
										if(match(fieldsToSave, secondaryChildNode.getNodeName(), sheet, nodesToLookTo[i])) {
											rowTitle.createCell(indice).setCellValue(secondaryChildNode.getNodeName());
											rowData.createCell(indice).setCellValue(secondaryChildNode.getTextContent());
											indice++;
											System.out.println(secondaryChildNode.getNodeName()+" = "+secondaryChildNode.getTextContent());
										}
										
										secondaryChildNode = secondaryChildNode.getNextSibling();
									}
								}
								else {
									if(match(fieldsToSave, firstChildNode.getNodeName(), sheet, nodesToLookTo[i])) {
										rowTitle.createCell(indice).setCellValue(firstChildNode.getNodeName());
										rowData.createCell(indice).setCellValue(firstChildNode.getTextContent());
										indice++;
										System.out.println(firstChildNode.getNodeName()+" = "+firstChildNode.getTextContent());
									}
									
								}
								firstChildNode = firstChildNode.getNextSibling();
							}
							
						}
						else {
							
							while(firstChildNode != null) {
								
								if(firstChildNode.getFirstChild() != null && firstChildNode.getFirstChild().hasChildNodes()) {
									
									Node thirdChild = firstChildNode.getFirstChild();
									while(thirdChild != null) {
										if(match(fieldsToSave, thirdChild.getNodeName(), sheet, nodesToLookTo[i])) {
											rowTitle.createCell(indice).setCellValue(thirdChild.getNodeName());
											rowData.createCell(indice).setCellValue(thirdChild.getTextContent());
											indice++;
											System.out.println(thirdChild.getNodeName()+" = "+thirdChild.getTextContent());
										}
										
										while(thirdChild.getFirstChild() != null && thirdChild.getFirstChild().hasChildNodes()) {
											
											thirdChild = thirdChild.getFirstChild();
										}
										thirdChild = thirdChild.getNextSibling();
									}
								}
								else {
									if(match(fieldsToSave, firstChildNode.getNodeName(), sheet, nodesToLookTo[i])) {
										rowTitle.createCell(indice).setCellValue(firstChildNode.getNodeName());
										rowData.createCell(indice).setCellValue(firstChildNode.getTextContent());
										indice++;
										System.out.println(firstChildNode.getNodeName()+" = "+firstChildNode.getTextContent());
									}
									
								}
								firstChildNode = firstChildNode.getNextSibling();
							}
						}
					}
				}
			}
		}
		//String fieldsCompl[] = {"infCpl"};
		NodeList node = doc.getElementsByTagName("infCpl");
		String valGasolina = searchForBCComplementary(node, "Gasolina - B.Calc.");
		String valDiesel = searchForBCComplementary(node, "Diesel - B.Calc.");
		if(valGasolina != null || valDiesel != null) {
			
			rowTitle.createCell(indice).setCellValue("VAL. BC COMPLEMENTAR");
			sheet.autoSizeColumn(indice);
			if(qtdProd == 1) {
				if(valGasolina != null) {
					System.out.println("Val. Compl. Gasolina = "+valGasolina);
					rowData.createCell(indice).setCellValue(valGasolina);
					//indice++;
				}else {
					System.out.println("Val. Compl. Diesel = "+valDiesel);
					rowData.createCell(indice).setCellValue(valDiesel);
					//indice++;
				}
			}else if(qtdProd > 1) {
				String cod;
				for(int x = 0; x < qtdProd; x++) {
					cod = sheet.getRow(line-x).getCell(6).getStringCellValue();
					if(cod.equals("15180002") || cod.equals("24311801") || cod.equals("15190002") || cod.equals("24146801")||
							cod.equals("24311801")) {//códigos de diesel
						System.out.println("Val. Compl. Diesel = "+valDiesel);
						sheet.getRow(line-x).createCell(indice).setCellValue(valDiesel);
					}
					else if(cod.equals("11110000") || cod.equals("11120000") || cod.equals("22149801") || cod.equals("22150801")){//códigos de gasolina
						System.out.println("Val. Compl. Gasolina = "+valGasolina);
						sheet.getRow(line-x).createCell(indice).setCellValue(valGasolina);
					}
				}
			}
			
		}
		
	}
	
	public static void main(String[] args) throws IOException, ParserConfigurationException, SAXException, EncryptedDocumentException, InvalidFormatException{
		// TODO Auto-generated method stub
		Workbook wb = new XSSFWorkbook();

		OutputStream fileOut = new FileOutputStream("/treinamento_cloudged/posto INOUYE 02-2014.xlsx");
		Sheet sheet = wb.createSheet("New Sheet");
		Row rowTitle = sheet.createRow(0);
		File arquivos[];
		File diretorio = new File("C:\\Users\\micro\\Desktop\\drive-download-20180820T162001Z-001\\2014\\02 - FEVEREIRO");
		arquivos = diretorio.listFiles();
		String extensionToFind = ".xml";
		ArrayList<String> titles = new ArrayList<String>();
		boolean signal = false;
		
		for(int i = 0; i < arquivos.length; i++) {
			int fileSize = (int) arquivos[i].getName().length();

			String nameFile = arquivos[i].getName();
			String extensao = nameFile.substring(fileSize-4, fileSize);
			if(extensao.compareToIgnoreCase(extensionToFind) == 0) {
				System.out.println("--------------"+arquivos[i].getName()+"--------------");
				indice = 0;
				inicialPos = 0;
				qtdProd = 0;
				ConvertXMLToXLSX(diretorio.getAbsolutePath()+"\\\\"+ arquivos[i].getName(),fileOut, sheet, rowTitle, titles);
				line++;
				
				signal = true;
			}

		}
		if(signal == false) {
			System.out.println("Não há nenhum arquivo "+extensionToFind+" nesse diretório;");
		}
		for(int l = 0; l < indice; l++) {
			sheet.autoSizeColumn(l);
		}
	
		wb.write(fileOut);
		wb.close();
		fileOut.close();	
	}


}