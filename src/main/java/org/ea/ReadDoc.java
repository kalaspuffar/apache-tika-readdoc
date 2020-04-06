package org.ea;

import org.apache.tika.Tika;
import org.apache.tika.extractor.EmbeddedDocumentExtractor;
import org.apache.tika.metadata.Metadata;
import org.apache.tika.parser.AutoDetectParser;
import org.apache.tika.parser.ParseContext;
import org.apache.tika.parser.Parser;
import org.apache.tika.sax.BasicContentHandlerFactory;
import org.xml.sax.ContentHandler;

import java.io.*;

public class ReadDoc {
    public static String parseDoc(String filename) {
        try {
            BasicContentHandlerFactory basicHandlerFactory = new BasicContentHandlerFactory(
                    BasicContentHandlerFactory.HANDLER_TYPE.XML, -1
            );
            AutoDetectParser parser = new AutoDetectParser();
            ContentHandler handler = basicHandlerFactory.getNewContentHandler();
            Metadata metadata = new Metadata();

            Tika tika = new Tika();
            ParseContext context = new ParseContext();
            context.set(
                EmbeddedDocumentExtractor.class,
                new FileEmbeddedDocumentExtractor(tika.getDetector())
            );
            context.set(Parser.class, parser);

            File docFile = new File(filename);
            InputStream stream = new FileInputStream(docFile);
            parser.parse(stream, handler, metadata, context);

            return handler.toString();
        } catch (Exception e) {
            e.printStackTrace();
        }
        return null;
    }

    public static void main(String[] args) {
        if(args.length < 1) {
            System.out.println("java -jar ReadDoc.jar [word-doc]");
            System.exit(0);
        }
        deleteDir(new File("work"));

        String docHtml = parseDoc(args[0]);
        System.out.println(docHtml);
    }

    public static void deleteDir(File file) {
        File[] contents = file.listFiles();
        if (contents != null) {
            for (File f : contents) {
                deleteDir(f);
            }
        }
        file.delete();
    }
}
