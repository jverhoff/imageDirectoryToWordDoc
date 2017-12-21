/*
John Verhoff

12/20/2017

insert all images from directory into word doc in same directory
 */
import java.io.*;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.*;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import javax.swing.*;
import javax.swing.filechooser.FileSystemView;

public class ImagesWord {
    public static void main(String[] args) throws IOException{
        XWPFDocument docx = new XWPFDocument();
        XWPFParagraph par = docx.createParagraph();
        XWPFRun run = par.createRun();

        JFileChooser jfc = new JFileChooser(FileSystemView.getFileSystemView().getDefaultDirectory());
        jfc.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
        jfc.showOpenDialog(null);
        String name = jfc.getSelectedFile().toString();

        final File folder = new File(name);

        FileOutputStream out = new FileOutputStream(name + "\\newWordDocImages.docx");
        int format = XWPFDocument.PICTURE_TYPE_JPEG;

        int i = 0;
        for(final File fileEntry : folder.listFiles()){
            if(fileEntry.toString().toLowerCase().endsWith(".jpg") || fileEntry.toString().toLowerCase().endsWith(".png")) {
                i++;
                String imgFile = fileEntry.toString();
                try {
                    run.addPicture(new FileInputStream(imgFile), format, imgFile, Units.toEMU(412), Units.toEMU(300)); //
                    // 200x200
                } catch (InvalidFormatException e) {
                    e.printStackTrace();
                }
                run.addBreak(BreakType.TEXT_WRAPPING);
                run.setText("Photograph " + i + " - ");
                run.addBreak();
                run.addBreak();
                run.setFontSize(12);
            }
            // pixels
        }
        docx.write(out);
        out.close();
    }
}
