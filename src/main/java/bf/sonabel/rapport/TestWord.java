package bf.sonabel.rapport;
import java.io.File;
import java.io.FileOutputStream;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
public class TestWord {



    public static void main(String[] args)throws Exception {

        //Création de document vide (en mémoire)
        XWPFDocument document = new XWPFDocument();

        //Création du document destination sur le disque
        FileOutputStream out = new FileOutputStream(new File("C:\\FormationJAVA\\Jour9\\createparagraph.docx"));

        //Création du paragraphe
        XWPFParagraph paragraph = document.createParagraph();
        XWPFRun run = paragraph.createRun();
        //Remplir le paragraphe
        run.setText("Bonjour Texte du paragraphe");

        //Ecrire document mémoire dans out physique
                document.write(out);
                //fermer le document
        out.close();
        System.out.println("createparagraph.docx written successfully");
    }
}



