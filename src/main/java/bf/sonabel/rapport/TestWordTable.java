package bf.sonabel.rapport;

import org.apache.poi.xwpf.usermodel.*;

import java.io.File;
import java.io.FileOutputStream;

public class TestWordTable {


    public static void main(String[] args) throws Exception {

        //Création de document vide (en mémoire)
        XWPFDocument document = new XWPFDocument();

        //Création du document destination sur le disque
        FileOutputStream out = new FileOutputStream(new File("C:\\FormationJAVA\\Jour9\\createparagraphTable.docx"));

        //Création du paragraphe
        XWPFParagraph paragraph = document.createParagraph();
        XWPFRun run = paragraph.createRun();
        //Remplir le paragraphe
        run.setText("Bonjour Texte du paragraphe suivi par un tableau ");

        //créer un tableau
        XWPFTable table = document.createTable();

        //créer les lignes
        // créer la première ligne
        XWPFTableRow tableRowOne = table.getRow(0);
        tableRowOne.getCell(0).setText("Nom");
        tableRowOne.addNewTableCell().setText("Prénom");
        tableRowOne.addNewTableCell().setText("Age");

        for(int i = 0; i < 10; i++) {
            XWPFTableRow tableRowTwo = table.createRow();
            for(int j = 0; j < 3 ; j++) {
                tableRowTwo.getCell(j).setText("Cellule " + i + " , " + j);

            }
        }

        //Ecrire document mémoire dans out physique
        document.write(out);
        //fermer le document
        out.close();
        for(int i = 0; i < 10; i++) {

            for(int j = 0; j < 3 ; j++) {
                System.out.print("Cellule " + i + " , " + j +" | ");

            }
            System.out.println();
        }
    }
}



