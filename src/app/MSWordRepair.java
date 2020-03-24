package app;

import java.util.ArrayList;
import java.util.List;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.xmlbeans.XmlCursor;
import org.w3c.dom.Node;

/**
 * This class provides static methods for repairing some of the faulty cases
 * which can occur with bookmarks in MS Word files.<br>
 * These include:<br>
 * - bookmarks nested<br>
 * - bookmark start and end on different xml levels<br>
 * - bookmark start and end in different paragraphs<br>
 * - bookmark end missing completely<br>
 */
public class MSWordRepair {
  /**
   * Repairs a couple of possible problems in the given document, that may disturb
   * further processing of the file.<br>
   * <br>
   * 
   * @param doc the document to be repaired
   */
  public final static void repair(XWPFDocument doc) {
    // repair bookmarkend position
    boolean startOverWithRepair = false;
    do {
      startOverWithRepair = false;
      List<Node> bookmarks = collectAllBookmarkStarts(doc);

      for (int i = 0; i < bookmarks.size(); i++) {
        // the parent of each bookmark must also contain the (correct) bookmarkend
        // element
        // if not, we need to find the bookmarkend and
        //
        // - if it is on the same level as the start element, we delete all sibling
        // branches from the mutual
        // parent
        // on and merge the two left over branches into one
        //
        // - if it is on a different level than the start element, we move the
        // bookmarkend tag to the end of the
        // start's parent
        //
        // - if we cannot find it, we make up a bookmarkend tag at the end of the
        // start's parent
        //
        // HINT: if we delete other bookmarks inbetween then that's how it is. We need
        // to catch when a bookmark
        // has
        // gotten invalid.
        //

        boolean bookmarkIsOk = false;
        Node bookmark = bookmarks.get(i);
        String startId = DOMHelpers.getIdFromNode(bookmark);

        Node next = bookmark.getNextSibling();
        while (next != null) {
          if (DOMHelpers.isBookmarkEnd(next)) {
            String endId = DOMHelpers.getIdFromNode(next);
            if (endId.equals(startId)) {
              bookmarkIsOk = true;
              break;
            }
          }

          if (bookmarkIsOk) {
            break;
          }
          next = next.getNextSibling();
        }

        if (!bookmarkIsOk) {
          List<Node> bmEnds = collectAllBookmarkEnds(doc);
          Node bmEnd = null;
          for (int j = 0; j < bmEnds.size(); j++) {
            bmEnd = bmEnds.get(j);
            String endId = DOMHelpers.getIdFromNode(bmEnd);

            if (endId.equals(startId)) {
              break;
            }
            bmEnd = null;
          }

          if (bmEnd != null) {
            ArrayList<Node> parentsStart = new ArrayList<Node>();
            ArrayList<Node> parentsEnd = new ArrayList<Node>();
            Node currNode = bookmark.getParentNode();
            while (currNode != null) {
              parentsStart.add(0, currNode);
              currNode = currNode.getParentNode();
            }

            currNode = bmEnd.getParentNode();
            while (currNode != null) {
              parentsEnd.add(0, currNode);
              currNode = currNode.getParentNode();
            }

            int startLevel = parentsStart.size();
            int endLevel = parentsEnd.size();

            int parentLevel = 0;
            int min = Math.min(parentsStart.size(), parentsEnd.size());
            Node mutParent = null;
            Node endParentChild = null;
            Node startParentChild = null;
            if (startLevel == endLevel) {
              // levels are identical, delete everything between

              for (parentLevel = 0; parentLevel < min; parentLevel++) {
                startParentChild = parentsStart.get(parentLevel);
                endParentChild = parentsEnd.get(parentLevel);

                if (parentsStart.get(parentLevel) != parentsEnd.get(parentLevel)) {
                  parentLevel--;
                  mutParent = parentsStart.get(parentLevel);
                  break;
                }
              }

              currNode = startParentChild.getNextSibling();
              while (currNode != endParentChild) {
                mutParent.removeChild(currNode);
                currNode = startParentChild.getNextSibling();
              }

              currNode = endParentChild.getFirstChild();
              while (currNode != null) {
                startParentChild.appendChild(currNode);
                currNode = endParentChild.getFirstChild();
              }

              mutParent.removeChild(endParentChild);
            } else {
              for (parentLevel = 0; parentLevel < min; parentLevel++) {
                endParentChild = parentsEnd.get(parentLevel);

                if (parentsStart.get(parentLevel) != parentsEnd.get(parentLevel)) {
                  parentLevel--;
                  mutParent = parentsStart.get(parentLevel);
                  break;
                }
              }

              if (mutParent == null) {
                mutParent = parentsStart.get(min - 1);
              }

              startParentChild = mutParent.getFirstChild();
              while (!parentsStart.contains(startParentChild) && bookmark != startParentChild) {
                startParentChild = startParentChild.getNextSibling();
              }

              currNode = startParentChild.getNextSibling();
              while (currNode != null && currNode != endParentChild && currNode != bmEnd
                  && !parentsEnd.contains(currNode)) {
                mutParent.removeChild(currNode);

                currNode = startParentChild.getNextSibling();
              }

              bookmark.getParentNode().appendChild(bmEnd);
            }
          } else {
            Node newEnd = DOMHelpers.createBookmarkEnd(bookmark);
            DOMHelpers.insertNodeAfter(newEnd, bookmark);
          }
          startOverWithRepair = true;
        }

        if (startOverWithRepair) {
          break;
        }
      }
    } while (startOverWithRepair);

    // repair nestings
    do {
      startOverWithRepair = false;

      List<Node> bmStartAndEnds = collectBMStartAndEndsInSequence(doc);
      boolean expect = false;
      boolean unexpected = false;
      String expected = "";
      Node bmStartNode = null;

      for (Node node : bmStartAndEnds) {
        if (!expect) {
          expect = true;
          expected = DOMHelpers.getIdFromNode(node);
          bmStartNode = node;
        } else {
          expect = false;
          if (!expected.equals(DOMHelpers.getIdFromNode(node))) {
            unexpected = true;
            break;
          }
        }
      }

      if (unexpected) {
        // sequence of starts and ends is wrong -> try to repair
        // bmStartNode contains the last correct node (start)
        // search for end node with bmStartNode's id and move it right next to
        // bmStartNode
        String startId = DOMHelpers.getIdFromNode(bmStartNode);
        if (!startId.isEmpty()) {
          List<Node> ends = collectAllBookmarkEnds(doc);
          for (Node bmEndNode : ends) {
            if (DOMHelpers.getIdFromNode(bmEndNode).equals(startId)) {
              DOMHelpers.insertNodeAfter(bmEndNode, bmStartNode);
              startOverWithRepair = true;
            }
          }
        } else {
          // repair failed, document seems to be too corrupted to repair (for us)
        }
      }
    } while (startOverWithRepair);
  }

  /**
   * Iterates through the XML of the given document from start to end and collects
   * and returns bookmark start and bookmark end nodes in the order as they
   * appear. If then the same node id does not appear two times after another in
   * this list, the sequence of the bookmark borders is disturbed and must be
   * repaired.<br>
   * <br>
   * 
   * @param doc document to scan
   * @return list of bookmark starts and bookmark ends in the order of appearance
   * 
   */
  private final static List<Node> collectBMStartAndEndsInSequence(XWPFDocument doc) {
    List<String> listIds = new ArrayList<String>();
    List<Node> result = new ArrayList<Node>();

    XmlCursor c = doc.getDocument().newCursor();

    boolean repeat = true;

    while (repeat) {
      Node node = c.getDomNode();
      if (DOMHelpers.isBookmarkStart(node) || DOMHelpers.isBookmarkEnd(node)) {
        String listId = node.getLocalName();
        if (node.getAttributes().getLength() > 0) {
          listId += node.getAttributes().item(0).getNodeValue();
        }
        if (!listIds.contains(listId)) {
          result.add(node);
          listIds.add(listId);
        }
      }

      boolean succ = c.toFirstChild();
      if (!succ) {
        succ = c.toNextSibling();
      }

      while (!succ && repeat) {
        succ = c.toParent();
        if (!succ) {
          repeat = false;
          break;
        }

        succ = c.toNextSibling();
      }
    }

    return result;
  }

  /**
   * Iterates through the XML DOM tree of the given document and collects and
   * returns all nodes which represent bookmark starts.<br>
   * <br>
   * 
   * @param doc document to scan
   * @return list of bookmark starts
   * 
   */
  private final static List<Node> collectAllBookmarkStarts(XWPFDocument doc) {
    Node root = doc.getDocument().getDomNode();
    List<Node> nodes = DOMHelpers.collectAllNodes(root, DOMHelpers.NODE_BM_START);

    for (int i = 0; i < nodes.size(); i++) {
      Node node = nodes.get(i);
      Node nameNode = node.getAttributes().getNamedItem("w:name");
      String name = nameNode.getNodeValue();
      if (name != null && (name.toLowerCase().startsWith("_ref") || name.toLowerCase().startsWith("_toc"))) {
        nodes.remove(i--);
      }
    }

    return nodes;
  }

  /**
   * Iterates through the XML DOM tree of the given document and collects and
   * returns all nodes which represent bookmark ends.<br>
   * <br>
   * 
   * @param doc document to scan
   * @return list of bookmark ends
   * 
   */
  private final static List<Node> collectAllBookmarkEnds(XWPFDocument doc) {
    Node root = doc.getDocument().getDomNode();
    return DOMHelpers.collectAllNodes(root, DOMHelpers.NODE_BM_END);
  }
}
