package org.ea;

import org.apache.poi.poifs.filesystem.DirectoryEntry;
import org.apache.poi.poifs.filesystem.DocumentEntry;
import org.apache.poi.poifs.filesystem.DocumentInputStream;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.tika.config.TikaConfig;
import org.apache.tika.detect.Detector;
import org.apache.tika.extractor.EmbeddedDocumentExtractor;
import org.apache.tika.io.FilenameUtils;
import org.apache.tika.io.IOUtils;
import org.apache.tika.io.TikaInputStream;
import org.apache.tika.metadata.Metadata;
import org.apache.tika.mime.MediaType;
import org.apache.tika.mime.MimeTypeException;
import org.xml.sax.ContentHandler;
import org.xml.sax.SAXException;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;

public class FileEmbeddedDocumentExtractor implements EmbeddedDocumentExtractor {
    private int count = 0;
    private final TikaConfig config = TikaConfig.getDefaultConfig();

    private static String extractDir = "work";
    private static Detector detector;

    public FileEmbeddedDocumentExtractor(Detector detector) {
        FileEmbeddedDocumentExtractor.detector = detector;
    }

    public boolean shouldParseEmbedded(Metadata metadata) {
        return true;
    }

    public void parseEmbedded(InputStream inputStream, ContentHandler contentHandler, Metadata metadata, boolean outputHtml) throws SAXException, IOException {

        if (!inputStream.markSupported()) {
            inputStream = TikaInputStream.get(inputStream);
        }
        MediaType contentType = detector.detect(inputStream, metadata);

        String name = metadata.get("resourceName");
        File outputFile = null;
        if (name == null) {
            name = "file" + count++;
        }
        outputFile = getOutputFile(name, metadata, contentType);


        File parent = outputFile.getParentFile();
        if (!parent.exists()) {
            if (!parent.mkdirs()) {
                throw new IOException("unable to create directory \"" + parent + "\"");
            }
        }
        System.err.println("Extracting '"+name+"' ("+contentType+") to " + outputFile);

        try {
            FileOutputStream os = new FileOutputStream(outputFile);
            if (inputStream instanceof TikaInputStream) {
                TikaInputStream tin = (TikaInputStream) inputStream;

                if (tin.getOpenContainer() != null && tin.getOpenContainer() instanceof DirectoryEntry) {
                    POIFSFileSystem fs = new POIFSFileSystem();
                    copy((DirectoryEntry) tin.getOpenContainer(), fs.getRoot());
                    fs.writeFilesystem(os);
                } else {
                    IOUtils.copy(inputStream, os);
                }
            } else {
                IOUtils.copy(inputStream, os);
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private File getOutputFile(String name, Metadata metadata, MediaType contentType) {
        String ext = getExtension(contentType);
        if (name.indexOf('.')==-1 && contentType!=null) {
            name += ext;
        }

        name = name.replaceAll("\u0000", " ");
        String normalizedName = FilenameUtils.normalize(name);
        File outputFile = new File(extractDir, normalizedName);

        return outputFile;
    }

    private String getExtension(MediaType contentType) {
        String extReturn = ".bin";
        try {
            String ext = config.getMimeRepository().forName(
                    contentType.toString()).getExtension();
            if (ext != null) extReturn = ext;
        } catch (MimeTypeException e) {
            e.printStackTrace();
        }
        return extReturn;

    }

    protected void copy(DirectoryEntry sourceDir, DirectoryEntry destDir)
            throws IOException {
        for (org.apache.poi.poifs.filesystem.Entry entry : sourceDir) {
            if (entry instanceof DirectoryEntry) {
                // Need to recurse
                DirectoryEntry newDir = destDir.createDirectory(entry.getName());
                copy((DirectoryEntry) entry, newDir);
            } else {
                // Copy entry
                InputStream contents = null;
                try {
                    contents = new DocumentInputStream((DocumentEntry) entry);
                    destDir.createDocument(entry.getName(), contents);
                } catch (Exception e) {
                    e.printStackTrace();
                } finally {
                    contents.close();
                }
            }
        }
    }
}

