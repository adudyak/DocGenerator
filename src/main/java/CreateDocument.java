import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.w3c.dom.Document;
import org.w3c.dom.NodeList;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import java.io.File;
import java.io.FileOutputStream;

public class CreateDocument {
    public static void createDocFile() {
        try {

            File xmlFile = new File("\\Users\\AlexD\\IdeaProjects\\DocGenerator\\persons.xml"); //Get xml file
            DocumentBuilder dBuilder = DocumentBuilderFactory.newInstance().newDocumentBuilder();
            Document xml = dBuilder.parse(xmlFile);
            NodeList nodeList = xml.getElementsByTagName("firstLastName"); //Get list of nodes with First and Last names by tag

            for (int i = 0; i < nodeList.getLength(); i++) {
                XWPFDocument doc = new XWPFDocument();
                FileOutputStream fos = new FileOutputStream(new File("Приглашение " + nodeList.item(i).getTextContent() + ".doc")); //Create blank document

                XWPFParagraph tempParagraph = doc.createParagraph();
                XWPFRun tempRun = tempParagraph.createRun();
                tempRun.setText("GeeksForLess, Inc.");
                tempRun.setFontSize(14);
                tempParagraph.setAlignment(ParagraphAlignment.LEFT);

                tempParagraph = doc.createParagraph();
                tempRun = tempParagraph.createRun();
                tempRun.setFontSize(14);
                tempParagraph.setAlignment(ParagraphAlignment.LEFT);
                tempRun.setText("г. Николаев, ул. Космонавтов 130В");

                tempParagraph = doc.createParagraph();
                tempRun = tempParagraph.createRun();
                tempRun.setFontSize(16);
                tempParagraph.setAlignment(ParagraphAlignment.CENTER);
                tempRun.setText("Приглашение");

                tempParagraph = doc.createParagraph();
                tempRun = tempParagraph.createRun();
                tempRun.setFontSize(14);
                tempRun.setText("   Уважаемый " + nodeList.item(i).getTextContent() + "! Приглашаем Вас посетить нашу компанию и погладить Маньку.");

                tempParagraph = doc.createParagraph();
                tempRun = tempParagraph.createRun();
                tempRun.setFontSize(14);
                tempRun.setText("Руководитель отдела кадров             Надежда Маховик");

                doc.write(fos);
                fos.close();

                System.out.println("Документ \"Приглашение " + nodeList.item(i).getTextContent() + ".doc\"" + " успешно создан!");
            }

        } catch (Exception e) {
            System.out.println(e.getMessage());
        }
    }

    public static void main(String[] args) {

        createDocFile();

    }

}