import java.io.*;

import javax.xml.parsers.*;
import org.w3c.dom.*;
import org.xml.sax.InputSource;
import java.io.StringReader;

import javax.xml.transform.*;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
import java.io.StringWriter;

public class xmlMod {

public static Document parseXmlString(String xmlString) throws Exception {
    DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
    DocumentBuilder builder = factory.newDocumentBuilder();
    
    System.out.println("Document is fetched");
    return builder.parse(new InputSource(new StringReader(xmlString)));
}

public static void removeNode(Node n) {	
	Node parent = n.getParentNode();
    	if (parent != null) {
        parent.removeChild(n);
        System.out.println(">>> Removed node.");
    }
}

public static void modifyNode(Document doc, String nd) {
    // 1. Get a list of all elements with the tag name "item"
    NodeList itemList = doc.getElementsByTagName(nd);

	
	int Len = itemList.getLength();
	for(int i = 0; i < Len; i++) {
		Element el = (Element) itemList.item(i);
		System.out.println("=> " + el.getNodeName() + " :: " + el.getTextContent() + " -> " + el.getAttribute("state"));
		if (el.hasAttribute("state")) {
            		el.removeAttribute("state");
			System.out.println("Attribute removed.");
		}
	}

	/*for(int i = Len-1; i >= 0; i--) {
		Element el = (Element) itemList.item(i);
		if(el.getAttribute("state") != null) removeNode(itemList.item(i));
	}*/


    if (itemList.getLength() > 0) {
        // 2. Target the first <item> element
        Element item = (Element) itemList.item(0);

        // --- Modification Examples ---
        
        // A. Modify the content (text node)
        item.setTextContent("NewModifiedValue");

        // B. Modify an attribute
        item.setAttribute("id", "100");
        
        // C. Add a new child element
        /*Element newChild = doc.createElement("status");
        newChild.setTextContent("UPDATED");
        item.appendChild(newChild);*/
        
        System.out.println("Node successfully modified in memory.");
    }
}

public static String serializeXmlDocument(Document doc) throws Exception {
    TransformerFactory tf = TransformerFactory.newInstance();
    Transformer transformer = tf.newTransformer();
    
    // Optional: for pretty-printing the output
    transformer.setOutputProperty(OutputKeys.INDENT, "yes");
    transformer.setOutputProperty("{http://xml.apache.org/xslt}indent-amount", "4");

    StringWriter writer = new StringWriter();
    transformer.transform(new DOMSource(doc), new StreamResult(writer));
    
    return writer.getBuffer().toString();
}


    public static void main(String[] args) throws IOException {
    String xml = "<root><sheet>First</sheet><sheet state='hidden'>Second</sheet></root>";
    try {
        Document doc = parseXmlString(xml);
        modifyNode(doc, "sheet");
        String res = serializeXmlDocument(doc);
        System.out.println(res);
    } catch(Exception e) {}
    }
}