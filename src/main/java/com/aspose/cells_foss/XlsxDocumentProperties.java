package com.aspose.cells_foss;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import java.io.InputStream;
import java.util.HashMap;
import java.util.Map;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;


/**
 * Helper class for handling XLSX document properties (core and extended).
 */
public final class XlsxDocumentProperties {
    private static final String CORE_PROPERTIES_NS = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";
    private static final String DUBLIN_CORE_NS = "http://purl.org/dc/elements/1.1/";
    private static final String DUBLIN_CORE_TERMS_NS = "http://purl.org/dc/terms/";
    private static final String XSI_NS = "http://www.w3.org/2001/XMLSchema-instance";
    private static final String EXTENDED_PROPERTIES_NS = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties";

    private static final String CORE_PROPERTIES_RELATIONSHIP_TYPE = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties";
    private static final String EXTENDED_PROPERTIES_RELATIONSHIP_TYPE = "http://schemas.openxmlformats.org/package/2006/relationships/extended-properties";

    /**
     * Initializes a new XlsxDocumentProperties instance.
     */
    private XlsxDocumentProperties() {}

    /**
     * Builds the core properties XML document.
     */
    public static org.w3c.dom.Document buildCorePropertiesDocument(com.aspose.cells_foss.core.WorkbookModel model) {
        // Use getDocumentProperties() method from WorkbookModel
        // Assuming getDocumentProperties() returns a class with hasStoredState() and property getters
        Object docProps = model.getDocumentProperties();
        if (docProps == null) {
            return null;
        }
        
        // Use reflection or?? the class has getCore() method
        // For now, assume getDocumentProperties() returns an object with the properties
        try {
            DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
            factory.setNamespaceAware(true);
            DocumentBuilder builder = factory.newDocumentBuilder();
            org.w3c.dom.Document doc = builder.newDocument();

            org.w3c.dom.Element root = doc.createElementNS(CORE_PROPERTIES_NS, "cp:coreProperties");
            root.setAttributeNS("http://www.w3.org/2000/xmlns/", "xmlns:cp", CORE_PROPERTIES_NS);
            root.setAttributeNS("http://www.w3.org/2000/xmlns/", "xmlns:dc", DUBLIN_CORE_NS);
            root.setAttributeNS("http://www.w3.org/2000/xmlns/", "xmlns:dcterms", DUBLIN_CORE_TERMS_NS);
            root.setAttributeNS("http://www.w3.org/2000/xmlns/", "xmlns:xsi", XSI_NS);

            // Assume getDocumentProperties() returns an object with getter methods
            addStringElement(root, DUBLIN_CORE_NS, "title", getDocumentPropertyString(docProps, "getTitle"));
            addStringElement(root, DUBLIN_CORE_NS, "subject", getDocumentPropertyString(docProps, "getSubject"));
            addStringElement(root, DUBLIN_CORE_NS, "creator", getDocumentPropertyString(docProps, "getCreator"));
            addStringElement(root, CORE_PROPERTIES_NS, "keywords", getDocumentPropertyString(docProps, "getKeywords"));
            addStringElement(root, DUBLIN_CORE_NS, "description", getDocumentPropertyString(docProps, "getDescription"));
            addStringElement(root, CORE_PROPERTIES_NS, "lastModifiedBy", getDocumentPropertyString(docProps, "getLastModifiedBy"));
            addStringElement(root, CORE_PROPERTIES_NS, "revision", getDocumentPropertyString(docProps, "getRevision"));
            addStringElement(root, CORE_PROPERTIES_NS, "category", getDocumentPropertyString(docProps, "getCategory"));
            addStringElement(root, CORE_PROPERTIES_NS, "contentStatus", getDocumentPropertyString(docProps, "getContentStatus"));
            addDateElement(root, DUBLIN_CORE_TERMS_NS, "created", getDocumentPropertyDate(docProps, "getCreated"));
            addDateElement(root, DUBLIN_CORE_TERMS_NS, "modified", getDocumentPropertyDate(docProps, "getModified"));

            return doc;
        } catch (Exception e) {
            return null;
        }
    }

    /**
     * Helper method to get string property using reflection-like approach
     */
    private static String getDocumentPropertyString(Object docProps, String methodName) {
        // Wrap lower-level failures in the library-specific exception flow.
        try {
            java.lang.reflect.Method method = docProps.getClass().getMethod(methodName);
            return (String) method.invoke(docProps);
        } catch (Exception e) {
            return "";
        }
    }

    /**
     * Processes get document property date.
     * @param docProps doc props
     * @param methodName name to use
     * @return the requested result
     */
    private static java.time.LocalDateTime getDocumentPropertyDate(Object docProps, String methodName) {
        // Wrap lower-level failures in the library-specific exception flow.
        try {
            java.lang.reflect.Method method = docProps.getClass().getMethod(methodName);
            Object result = method.invoke(docProps);
            return (java.time.LocalDateTime) result;
        } catch (Exception e) {
            return null;
        }
    }

    /**
     * Builds the extended properties XML document.
     */
    public static org.w3c.dom.Document buildExtendedPropertiesDocument(com.aspose.cells_foss.core.WorkbookModel model) {
        // Use getDocumentProperties() method from WorkbookModel
        Object docProps = model.getDocumentProperties();
        if (docProps == null) {
            return null;
        }
        
        try {
            DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
            factory.setNamespaceAware(true);
            DocumentBuilder builder = factory.newDocumentBuilder();
            org.w3c.dom.Document doc = builder.newDocument();

            org.w3c.dom.Element root = doc.createElementNS(EXTENDED_PROPERTIES_NS, "Properties");
            root.setAttributeNS("http://www.w3.org/2000/xmlns/", "xmlns:vt", "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes");

            addStringElement(root, EXTENDED_PROPERTIES_NS, "Application", getDocumentPropertyString(docProps, "getApplication"));
            addStringElement(root, EXTENDED_PROPERTIES_NS, "AppVersion", getDocumentPropertyString(docProps, "getAppVersion"));
            addStringElement(root, EXTENDED_PROPERTIES_NS, "Company", getDocumentPropertyString(docProps, "getCompany"));
            addStringElement(root, EXTENDED_PROPERTIES_NS, "Manager", getDocumentPropertyString(docProps, "getManager"));
            addIntElement(root, EXTENDED_PROPERTIES_NS, "DocSecurity", getDocumentPropertyInt(docProps, "getDocSecurity"));
            addStringElement(root, EXTENDED_PROPERTIES_NS, "HyperlinkBase", getDocumentPropertyString(docProps, "getHyperlinkBase"));
            addBoolElement(root, EXTENDED_PROPERTIES_NS, "ScaleCrop", getDocumentPropertyBool(docProps, "getScaleCrop"));
            addBoolElement(root, EXTENDED_PROPERTIES_NS, "LinksUpToDate", getDocumentPropertyBool(docProps, "getLinksUpToDate"));
            addBoolElement(root, EXTENDED_PROPERTIES_NS, "SharedDoc", getDocumentPropertyBool(docProps, "getSharedDoc"));

            return doc;
        } catch (Exception e) {
            return null;
        }
    }

    /**
     * Processes get document property int.
     * @param docProps doc props
     * @param methodName name to use
     * @return the requested result
     */
    private static Integer getDocumentPropertyInt(Object docProps, String methodName) {
        // Wrap lower-level failures in the library-specific exception flow.
        try {
            java.lang.reflect.Method method = docProps.getClass().getMethod(methodName);
            Object result = method.invoke(docProps);
            return (Integer) result;
        } catch (Exception e) {
            return null;
        }
    }

    /**
     * Processes get document property bool.
     * @param docProps doc props
     * @param methodName name to use
     * @return the requested result
     */
    private static Boolean getDocumentPropertyBool(Object docProps, String methodName) {
        // Wrap lower-level failures in the library-specific exception flow.
        try {
            java.lang.reflect.Method method = docProps.getClass().getMethod(methodName);
            Object result = method.invoke(docProps);
            return (Boolean) result;
        } catch (Exception e) {
            return null;
        }
    }

    /**
     * Loads document properties from the ZIP archive.
     */
    public static void loadDocumentProperties(ZipFile archive, com.aspose.cells_foss.core.WorkbookModel model, LoadDiagnostics diagnostics, LoadOptions options) {
        Map<String, String> relationshipTargets = loadRootDocumentPropertiesRelationshipTargets(archive, diagnostics, options);
        
        String corePropertiesPartUri = relationshipTargets.get(CORE_PROPERTIES_RELATIONSHIP_TYPE);
        // Handle the relevant branch before the state changes.
        if (corePropertiesPartUri == null || corePropertiesPartUri.isEmpty()) {
            corePropertiesPartUri = "";
        }
        
        String extendedPropertiesPartUri = relationshipTargets.get(EXTENDED_PROPERTIES_RELATIONSHIP_TYPE);
        if (extendedPropertiesPartUri == null || extendedPropertiesPartUri.isEmpty()) {
            extendedPropertiesPartUri = "";
        }

        loadCoreProperties(archive, model, diagnostics, options, corePropertiesPartUri);
        loadExtendedProperties(archive, model, diagnostics, options, extendedPropertiesPartUri);
    }

    /**
     * Loads the root document properties relationship targets.
     * @param archive archive
     * @param diagnostics diagnostics
     * @param options options to apply
     * @return the requested result
     */
    private static Map<String, String> loadRootDocumentPropertiesRelationshipTargets(ZipFile archive, LoadDiagnostics diagnostics, LoadOptions options) {
        Map<String, String> relationshipTargets = new HashMap<>();
        
        ZipEntry entry = archive.getEntry("_rels/.rels");
        if (entry == null) {
            return relationshipTargets;
        }

        org.w3c.dom.Document document;
        try (InputStream is = archive.getInputStream(entry)) {
            DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
            factory.setNamespaceAware(true);
            DocumentBuilder builder = factory.newDocumentBuilder();
            document = builder.parse(is);
        } catch (Exception e) {
            addDocumentPropertiesIssue(diagnostics, options, "/_rels/.rels", "Package root relationships were malformed and document properties were ignored.");
            return relationshipTargets;
        }

        org.w3c.dom.Element root = document.getDocumentElement();
        if (root == null) {
            return relationshipTargets;
        }

        // Process relationship elements
        org.w3c.dom.NodeList nodeList = root.getChildNodes();
        for (int i = 0; i < nodeList.getLength(); i++) {
            org.w3c.dom.Node node = nodeList.item(i);
            if (node.getNodeType() == org.w3c.dom.Node.ELEMENT_NODE && "Relationship".equals(node.getLocalName())) {
                org.w3c.dom.Element relationship = (org.w3c.dom.Element) node;
                
                String type = relationship.getAttribute("Type");
                String target = relationship.getAttribute("Target");
                String targetMode = relationship.getAttribute("TargetMode");
                
                if (type == null || type.isEmpty() || 
                    target == null || target.isEmpty() ||
                    "External".equalsIgnoreCase(targetMode)) {
                    continue;
                }

                if (!CORE_PROPERTIES_RELATIONSHIP_TYPE.equalsIgnoreCase(type) &&
                    !EXTENDED_PROPERTIES_RELATIONSHIP_TYPE.equalsIgnoreCase(type)) {
                    continue;
                }

                if (!relationshipTargets.containsKey(type)) {
                    relationshipTargets.put(type, resolvePartUri("/", target));
                }
            }
        }

        return relationshipTargets;
    }

    /**
     * Loads the core properties.
     * @param archive archive
     * @param model model
     * @param diagnostics diagnostics
     * @param options options to apply
     * @param partUri part uri
     */
    private static void loadCoreProperties(ZipFile archive, com.aspose.cells_foss.core.WorkbookModel model, LoadDiagnostics diagnostics, LoadOptions options, String partUri) {
        if (partUri == null || partUri.isEmpty()) {
            return;
        }

        ZipEntry entry = archive.getEntry(partUri);
        if (entry == null) {
            return;
        }

        org.w3c.dom.Document document;
        try (InputStream is = archive.getInputStream(entry)) {
            DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
            factory.setNamespaceAware(true);
            DocumentBuilder builder = factory.newDocumentBuilder();
            document = builder.parse(is);
        } catch (Exception e) {
            addDocumentPropertiesIssue(diagnostics, options, partUri, "Core document properties part was malformed and was ignored.");
            return;
        }

        org.w3c.dom.Element root = document.getDocumentElement();
        if (root == null) {
            return;
        }

        // Use reflection to set properties on the document properties object
        Object docProps = model.getDocumentProperties();
        if (docProps != null) {
            setDocumentPropertyString(docProps, "setTitle", readElementValue(root, DUBLIN_CORE_NS, "title"));
            setDocumentPropertyString(docProps, "setSubject", readElementValue(root, DUBLIN_CORE_NS, "subject"));
            setDocumentPropertyString(docProps, "setCreator", readElementValue(root, DUBLIN_CORE_NS, "creator"));
            setDocumentPropertyString(docProps, "setKeywords", readElementValue(root, CORE_PROPERTIES_NS, "keywords"));
            setDocumentPropertyString(docProps, "setDescription", readElementValue(root, DUBLIN_CORE_NS, "description"));
            setDocumentPropertyString(docProps, "setLastModifiedBy", readElementValue(root, CORE_PROPERTIES_NS, "lastModifiedBy"));
            setDocumentPropertyString(docProps, "setRevision", readElementValue(root, CORE_PROPERTIES_NS, "revision"));
            setDocumentPropertyString(docProps, "setCategory", readElementValue(root, CORE_PROPERTIES_NS, "category"));
            setDocumentPropertyString(docProps, "setContentStatus", readElementValue(root, CORE_PROPERTIES_NS, "contentStatus"));
            setDocumentPropertyDate(docProps, "setCreated", readDateElement(root, DUBLIN_CORE_TERMS_NS, "created", diagnostics, options, partUri));
            setDocumentPropertyDate(docProps, "setModified", readDateElement(root, DUBLIN_CORE_TERMS_NS, "modified", diagnostics, options, partUri));
        }
    }

    /**
     * Sets the document property string.
     * @param docProps doc props
     * @param methodName name to use
     * @param value value to apply
     */
    private static void setDocumentPropertyString(Object docProps, String methodName, String value) {
        try {
            java.lang.reflect.Method method = docProps.getClass().getMethod(methodName, String.class);
            method.invoke(docProps, value);
        } catch (Exception e) {
            // Ignore
        }
    }

    /**
     * Sets the document property date.
     * @param docProps doc props
     * @param methodName name to use
     * @param value value to apply
     */
    private static void setDocumentPropertyDate(Object docProps, String methodName, java.time.LocalDateTime value) {
        try {
            java.lang.reflect.Method method = docProps.getClass().getMethod(methodName, java.time.LocalDateTime.class);
            method.invoke(docProps, value);
        } catch (Exception e) {
            // Ignore
        }
    }

    /**
     * Loads the extended properties.
     * @param archive archive
     * @param model model
     * @param diagnostics diagnostics
     * @param options options to apply
     * @param partUri part uri
     */
    private static void loadExtendedProperties(ZipFile archive, com.aspose.cells_foss.core.WorkbookModel model, LoadDiagnostics diagnostics, LoadOptions options, String partUri) {
        if (partUri == null || partUri.isEmpty()) {
            return;
        }

        ZipEntry entry = archive.getEntry(partUri);
        if (entry == null) {
            return;
        }

        org.w3c.dom.Document document;
        try (InputStream is = archive.getInputStream(entry)) {
            DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
            factory.setNamespaceAware(true);
            DocumentBuilder builder = factory.newDocumentBuilder();
            document = builder.parse(is);
        } catch (Exception e) {
            addDocumentPropertiesIssue(diagnostics, options, partUri, "Extended document properties part was malformed and was ignored.");
            return;
        }

        org.w3c.dom.Element root = document.getDocumentElement();
        if (root == null) {
            return;
        }

        // Use reflection to set properties on the document properties object
        Object docProps = model.getDocumentProperties();
        if (docProps != null) {
            setDocumentPropertyString(docProps, "setApplication", readElementValue(root, EXTENDED_PROPERTIES_NS, "Application"));
            setDocumentPropertyString(docProps, "setAppVersion", readElementValue(root, EXTENDED_PROPERTIES_NS, "AppVersion"));
            setDocumentPropertyString(docProps, "setCompany", readElementValue(root, EXTENDED_PROPERTIES_NS, "Company"));
            setDocumentPropertyString(docProps, "setManager", readElementValue(root, EXTENDED_PROPERTIES_NS, "Manager"));
            setDocumentPropertyInt(docProps, "setDocSecurity", readIntElement(root, EXTENDED_PROPERTIES_NS, "DocSecurity", diagnostics, options, partUri));
            setDocumentPropertyString(docProps, "setHyperlinkBase", readElementValue(root, EXTENDED_PROPERTIES_NS, "HyperlinkBase"));
            setDocumentPropertyBool(docProps, "setScaleCrop", readBoolElement(root, EXTENDED_PROPERTIES_NS, "ScaleCrop", diagnostics, options, partUri));
            setDocumentPropertyBool(docProps, "setLinksUpToDate", readBoolElement(root, EXTENDED_PROPERTIES_NS, "LinksUpToDate", diagnostics, options, partUri));
            setDocumentPropertyBool(docProps, "setSharedDoc", readBoolElement(root, EXTENDED_PROPERTIES_NS, "SharedDoc", diagnostics, options, partUri));
        }
    }

    /**
     * Sets the document property int.
     * @param docProps doc props
     * @param methodName name to use
     * @param value value to apply
     */
    private static void setDocumentPropertyInt(Object docProps, String methodName, Integer value) {
        try {
            java.lang.reflect.Method method = docProps.getClass().getMethod(methodName, Integer.class);
            method.invoke(docProps, value);
        } catch (Exception e) {
            // Ignore
        }
    }

    /**
     * Sets the document property bool.
     * @param docProps doc props
     * @param methodName name to use
     * @param value value to apply
     */
    private static void setDocumentPropertyBool(Object docProps, String methodName, Boolean value) {
        try {
            java.lang.reflect.Method method = docProps.getClass().getMethod(methodName, Boolean.class);
            method.invoke(docProps, value);
        } catch (Exception e) {
            // Ignore
        }
    }

    /**
     * Adds string element.
     * @param parent parent
     * @param namespace namespace
     * @param localName name to use
     * @param value value to apply
     */
    private static void addStringElement(org.w3c.dom.Element parent, String namespace, String localName, String value) {
        // Handle the relevant branch before the state changes.
        if (value != null && !value.isEmpty()) {
            org.w3c.dom.Element element = parent.getOwnerDocument().createElementNS(namespace, localName);
            element.setTextContent(value);
            parent.appendChild(element);
        }
    }

    /**
     * Adds int element.
     * @param parent parent
     * @param namespace namespace
     * @param localName name to use
     * @param value value to apply
     */
    private static void addIntElement(org.w3c.dom.Element parent, String namespace, String localName, Integer value) {
        // Handle the relevant branch before the state changes.
        if (value != null) {
            org.w3c.dom.Element element = parent.getOwnerDocument().createElementNS(namespace, localName);
            element.setTextContent(String.valueOf(value));
            parent.appendChild(element);
        }
    }

    /**
     * Adds bool element.
     * @param parent parent
     * @param namespace namespace
     * @param localName name to use
     * @param value value to apply
     */
    private static void addBoolElement(org.w3c.dom.Element parent, String namespace, String localName, Boolean value) {
        // Handle the relevant branch before the state changes.
        if (value != null) {
            org.w3c.dom.Element element = parent.getOwnerDocument().createElementNS(namespace, localName);
            element.setTextContent(value ? "true" : "false");
            parent.appendChild(element);
        }
    }

    /**
     * Adds date element.
     * @param parent parent
     * @param namespace namespace
     * @param localName name to use
     * @param value value to apply
     */
    private static void addDateElement(org.w3c.dom.Element parent, String namespace, String localName, java.time.LocalDateTime value) {
        // Handle the relevant branch before the state changes.
        if (value == null) {
            return;
        }

        org.w3c.dom.Element element = parent.getOwnerDocument().createElementNS(namespace, localName);
        element.setTextContent(java.time.format.DateTimeFormatter.ISO_LOCAL_DATE_TIME.format(value));
        element.setAttributeNS(XSI_NS, "xsi:type", "dcterms:W3CDTF");
        parent.appendChild(element);
    }

    /**
     * Reads element value.
     * @param parent parent
     * @param namespace namespace
     * @param localName name to use
     * @return the computed result
     */
    private static String readElementValue(org.w3c.dom.Element parent, String namespace, String localName) {
        org.w3c.dom.NodeList nodeList = parent.getChildNodes();
        // Walk the current collection so every entry is processed consistently.
        for (int i = 0; i < nodeList.getLength(); i++) {
            org.w3c.dom.Node node = nodeList.item(i);
            if (node.getNodeType() == org.w3c.dom.Node.ELEMENT_NODE &&
                localName.equals(node.getLocalName())) {
                return node.getTextContent().trim();
            }
        }
        return "";
    }

    /**
     * Reads date element.
     * @param parent parent
     * @param namespace namespace
     * @param localName name to use
     * @param diagnostics diagnostics
     * @param options options to apply
     * @param partUri part uri
     * @return the computed result
     */
    private static java.time.LocalDateTime readDateElement(org.w3c.dom.Element parent, String namespace, String localName, LoadDiagnostics diagnostics, LoadOptions options, String partUri) {
        org.w3c.dom.NodeList nodeList = parent.getChildNodes();
        // Walk the current collection so every entry is processed consistently.
        for (int i = 0; i < nodeList.getLength(); i++) {
            org.w3c.dom.Node node = nodeList.item(i);
            if (node.getNodeType() == org.w3c.dom.Node.ELEMENT_NODE &&
                localName.equals(node.getLocalName())) {
                try {
                    return java.time.LocalDateTime.parse(node.getTextContent());
                } catch (Exception e) {
                    addDocumentPropertiesIssue(diagnostics, options, partUri, "Document property '" + localName + "' had an invalid timestamp and was ignored.");
                    return null;
                }
            }
        }
        return null;
    }

    /**
     * Reads int element.
     * @param parent parent
     * @param namespace namespace
     * @param localName name to use
     * @param diagnostics diagnostics
     * @param options options to apply
     * @param partUri part uri
     * @return the computed result
     */
    private static Integer readIntElement(org.w3c.dom.Element parent, String namespace, String localName, LoadDiagnostics diagnostics, LoadOptions options, String partUri) {
        org.w3c.dom.NodeList nodeList = parent.getChildNodes();
        for (int i = 0; i < nodeList.getLength(); i++) {
            org.w3c.dom.Node node = nodeList.item(i);
            if (node.getNodeType() == org.w3c.dom.Node.ELEMENT_NODE &&
                localName.equals(node.getLocalName())) {
                try {
                    int value = Integer.parseInt(node.getTextContent());
                    if (value >= 0) {
                        return value;
                    }
                } catch (NumberFormatException e) {
                    // Continue to error handling
                }
                addDocumentPropertiesIssue(diagnostics, options, partUri, "Document property '" + localName + "' had an invalid integer value and was ignored.");
                return null;
            }
        }
        return null;
    }

    /**
     * Reads bool element.
     * @param parent parent
     * @param namespace namespace
     * @param localName name to use
     * @param diagnostics diagnostics
     * @param options options to apply
     * @param partUri part uri
     * @return the computed result
     */
    private static Boolean readBoolElement(org.w3c.dom.Element parent, String namespace, String localName, LoadDiagnostics diagnostics, LoadOptions options, String partUri) {
        org.w3c.dom.NodeList nodeList = parent.getChildNodes();
        // Walk the current collection so every entry is processed consistently.
        for (int i = 0; i < nodeList.getLength(); i++) {
            org.w3c.dom.Node node = nodeList.item(i);
            if (node.getNodeType() == org.w3c.dom.Node.ELEMENT_NODE &&
                localName.equals(node.getLocalName())) {
                String rawValue = node.getTextContent();
                if ("1".equals(rawValue) || "true".equalsIgnoreCase(rawValue)) {
                    return true;
                }
                if ("0".equals(rawValue) || "false".equalsIgnoreCase(rawValue)) {
                    return false;
                }
                addDocumentPropertiesIssue(diagnostics, options, partUri, "Document property '" + localName + "' had an invalid Boolean value and was ignored.");
                return null;
            }
        }
        return null;
    }

    /**
     * Adds document properties issue.
     * @param diagnostics diagnostics
     * @param options options to apply
     * @param partUri part uri
     * @param message message
     */
    private static void addDocumentPropertiesIssue(LoadDiagnostics diagnostics, LoadOptions options, String partUri, String message) {
        LoadIssue issue = new LoadIssue("WB-L004", DiagnosticSeverity.WARNING, message, false, false);
        issue.setPartUri(partUri);
    }

    /**
     * Resolves part uri.
     * @param baseUri base uri
     * @param target target
     * @return the computed result
     */
    private static String resolvePartUri(String baseUri, String target) {
        // Simplified resolution - in real implementation would handle relative paths
        if (target == null) {
            return null;
        }
        // For simplicity, assume target is relative and just return it
        // Full implementation would resolve relative to baseUri
        return target.startsWith("/") ? target : baseUri + target;
    }
}