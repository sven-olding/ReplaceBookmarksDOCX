package app;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.Arrays;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.xwpf.usermodel.BodyElementType;
import org.apache.poi.xwpf.usermodel.IBodyElement;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFFooter;
import org.apache.poi.xwpf.usermodel.XWPFHeader;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.xmlbeans.XmlCursor;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTP;
import org.w3c.dom.Element;
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

      replaceBookmarkContent(wordDoc, replaceList);

      wordDoc.write(fos);
    } catch (Exception e) {
      e.printStackTrace();
    }
  }

  private static void replaceBookmarkContent(XWPFDocument wordDoc, Map<String, String> replacementList) {
    String NEWLINE = System.getProperty("line.separator");
    XmlCursor cursor = null;

    MSWordRepair.repair(wordDoc);

    // collect all bookmark starts in document
    Node root = wordDoc.getDocument().getDomNode();
    List<Node> bookmarks = DOMHelpers.collectAllNodes(root, DOMHelpers.NODE_BM_START);

    // collect all bookmark starts in all page headers
    List<XWPFHeader> headers = wordDoc.getHeaderList();
    for (XWPFHeader header : headers) {
      List<IBodyElement> bodyElements = header.getBodyElements();
      if (bodyElements.size() > 0) {
        IBodyElement p = bodyElements.get(0);
        if (p.getElementType() == BodyElementType.PARAGRAPH) {
          XWPFParagraph x = (XWPFParagraph) p;
          CTP ctp = x.getCTP();
          cursor = ctp.newCursor();
        } else if (p.getElementType() == BodyElementType.TABLE) {
          XWPFTable x = (XWPFTable) p;
          cursor = x.getCTTbl().newCursor();
        }

        while (cursor.toParent() || cursor.toPrevSibling()) {
          // doit
        }
        bookmarks.addAll(DOMHelpers.collectAllNodes(cursor.getDomNode(), DOMHelpers.NODE_BM_START));
      }
    }

    // collect all bookmark starts in all page footers
    List<XWPFFooter> footers = wordDoc.getFooterList();
    for (XWPFFooter footer : footers) {
      List<IBodyElement> bodyElements = footer.getBodyElements();
      if (bodyElements.size() > 0) {
        IBodyElement p = bodyElements.get(0);
        if (p.getElementType() == BodyElementType.PARAGRAPH) {
          XWPFParagraph x = (XWPFParagraph) p;
          CTP ctp = x.getCTP();
          cursor = ctp.newCursor();
        } else if (p.getElementType() == BodyElementType.TABLE) {
          XWPFTable x = (XWPFTable) p;
          cursor = x.getCTTbl().newCursor();
        }

        while (cursor.toParent() || cursor.toPrevSibling()) {
          // doit
        }
        bookmarks.addAll(DOMHelpers.collectAllNodes(cursor.getDomNode(), DOMHelpers.NODE_BM_START));
      }
    }

    // start replacing
    for (Node start : bookmarks) {
      String name = DOMHelpers.getNameFromNode(start).toLowerCase();
      String value = replacementList.get(name);

      System.out.println(name + "=" + value);

      if (value != null) {
        Element parent = (Element) start.getParentNode();
        if (parent == null) {
          continue;
        }

        Node nextNode = start.getNextSibling();
        while (!DOMHelpers.isBookmarkEnd(nextNode)) {
          parent.removeChild(nextNode);
          nextNode = start.getNextSibling();
        }

        Node newRange = null;
        if (value.contains(NEWLINE)) {
          String[] values = value.split(NEWLINE);
          newRange = DOMHelpers.createRangeWithText(Arrays.asList(values), parent);
        } else if (value.contains("\n")) {
          String[] values = value.split("\n");
          newRange = DOMHelpers.createRangeWithText(Arrays.asList(values), parent);
        } else {
          newRange = DOMHelpers.createRangeWithText(value, parent);
        }

        Node rPrNode = DOMHelpers.clonePreviousRPr(start);
        if (rPrNode != null) {
          newRange.insertBefore(rPrNode, newRange.getFirstChild());
        }
        parent.insertBefore(newRange, nextNode);
      }
    }
  }
}