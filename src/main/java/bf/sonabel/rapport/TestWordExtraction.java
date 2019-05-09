package bf.sonabel.rapport;

import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import java.io.FileInputStream;

public class TestWordExtraction {

    public static void main(String[] args)throws Exception {

        XWPFDocument docx = new XWPFDocument(new FileInputStream("C:\\FormationJAVA\\Jour9\\createparagraph.docx"));

        //
        XWPFWordExtractor we = new XWPFWordExtractor(docx);
        System.out.println(we.getText());
        System.out.println("Deuxième document");
        docx = new XWPFDocument(new FileInputStream("C:\\FormationJAVA\\Jour9\\createparagraphTable.docx"));

        //
        we = new XWPFWordExtractor(docx);
        System.out.println(we.getText());
    }
}
