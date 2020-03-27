package app;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTBookmark;
import org.w3c.dom.Node;

public class App {
  public static void main(String[] args) throws Exception {
    try (FileInputStream fis = new FileInputStream(new File("./docx/Document.docx"));
        FileOutputStream fos = new FileOutputStream(new File("./docx/Result.docx"));) {
      XWPFDocument wordDoc = new XWPFDocument(fis);
      HashMap<String, String> replaceList = new HashMap<String, String>();
      replaceList.put("abnahmestelle", "Meine Abnahmestelle");
      replaceList.put("marktlokation", "MaLo");
      replaceList.put("preis", "1");
      replaceList.put("lieferzeitraumvon", "von");
      replaceList.put("lieferzeitraumbis", "bis");
      replaceList.put("vertragskonto", "12345");
      replaceList.put("einheit", "kWh");

      replaceBookmarkContent(wordDoc, replaceList);

      wordDoc.write(fos);
    } catch (Exception e) {
      e.printStackTrace();
    }
  }

  private static void replaceBookmarkContent(XWPFDocument wordDoc, Map<String, String> replacementList) {
    MSWordRepair.repair(wordDoc);
    stripUnneededBookmarks(wordDoc);
    for (XWPFParagraph p : wordDoc.getParagraphs()) {
      replaceBookmarkContent(p, replacementList);
    }
    for (XWPFTable t : wordDoc.getTables()) {
      for (XWPFTableRow row : t.getRows()) {
        for (XWPFTableCell cell : row.getTableCells()) {
          for (XWPFParagraph p : cell.getParagraphs()) {
            replaceBookmarkContent(p, replacementList);
          }
        }
      }
    }
  }

  private static void replaceBookmarkContent(XWPFParagraph paragraph, Map<String, String> replacementList) {
    for (CTBookmark bookmark : paragraph.getCTP().getBookmarkStartList()) {
      for (String key : replacementList.keySet()) {
        if (bookmark.getName().equalsIgnoreCase(key)) {
          Node nextNode = bookmark.getDomNode().getNextSibling();
          while (!(nextNode.getNodeName().contains("bookmarkEnd"))) {
            paragraph.getCTP().getDomNode().removeChild(nextNode);
            nextNode = bookmark.getDomNode().getNextSibling();
          }

          XWPFRun run = paragraph.createRun();
          run.setText(replacementList.get(key));

          Node styles = DOMHelpers.clonePreviousRPr(bookmark.getDomNode());
          Node runNode = run.getCTR().getDomNode();
          runNode.insertBefore(styles, runNode.getFirstChild());
          paragraph.getCTP().getDomNode().insertBefore(runNode, nextNode);
        }
      }
    }
  }

  private static void stripUnneededBookmarks(XWPFDocument wordDoc) {
    List<Node> startNodes = DOMHelpers.collectAllNodes(wordDoc.getDocument().getDomNode(), DOMHelpers.NODE_BM_START);
    List<Node> endNodes = DOMHelpers.collectAllNodes(wordDoc.getDocument().getDomNode(), DOMHelpers.NODE_BM_END);

    for (Node start : startNodes) {
      String bmName = DOMHelpers.getNameFromNode(start);
      if (bmName.equalsIgnoreCase("_GoBack")) {
        String startId = DOMHelpers.getIdFromNode(start);
        for (Node end : endNodes) {
          if (DOMHelpers.getIdFromNode(end).equals(startId)) {
            end.getParentNode().removeChild(end);
          }
        }
        start.getParentNode().removeChild(start);
      }
    }
  }
}