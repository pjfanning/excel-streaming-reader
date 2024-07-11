package com.github.pjfanning.xlsx.impl;

import com.github.pjfanning.poi.xssf.streaming.CommentsTableBase;
import com.github.pjfanning.poi.xssf.streaming.MapBackedCommentsTable;
import com.github.pjfanning.poi.xssf.streaming.MapBackedSharedStringsTable;
import com.github.pjfanning.poi.xssf.streaming.TempFileCommentsTable;
import com.github.pjfanning.poi.xssf.streaming.TempFileSharedStringsTable;
import com.github.pjfanning.xlsx.StreamingReader;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xssf.model.Comments;
import org.apache.poi.xssf.model.SharedStringsTable;

import java.io.IOException;
import java.io.InputStream;

/**
 * Keeps code that relies on poi-shared-strings out of the main codebase.
 */
public final class PoiSharedStringsSupport {
    public static Comments createTempFileCommentsTable(final StreamingReader.Builder builder) throws IOException {
        return new TempFileCommentsTable(
                builder.encryptCommentsTempFile(),
                builder.fullFormatRichText());
    }

    public static Comments createMapBackedCommentsTable(final StreamingReader.Builder builder) {
        return new MapBackedCommentsTable(builder.fullFormatRichText());
    }

    public static SharedStringsTable createTempFileSharedStringsTable(
            final StreamingReader.Builder builder) throws IOException {
        return new TempFileSharedStringsTable(
                builder.encryptSstTempFile(),
                builder.fullFormatRichText());
    }

    public static SharedStringsTable createTempFileSharedStringsTable(
            final OPCPackage pkg, final StreamingReader.Builder builder) throws IOException {
        return new TempFileSharedStringsTable(
                pkg,
                builder.encryptSstTempFile(),
                builder.fullFormatRichText());
    }

    public static SharedStringsTable createMapBackedSharedStringsTable(
            final StreamingReader.Builder builder) {
        return new MapBackedSharedStringsTable(builder.fullFormatRichText());
    }

    public static SharedStringsTable createMapBackedSharedStringsTable(
            final OPCPackage pkg, final StreamingReader.Builder builder) throws IOException {
        return new MapBackedSharedStringsTable(
                pkg,
                builder.fullFormatRichText());
    }

    public static void readComments(final Comments comments, final InputStream inputStream) throws IOException {
        if (comments instanceof CommentsTableBase) {
            ((CommentsTableBase) comments).readFrom(inputStream);
        }
    }

    private PoiSharedStringsSupport() {
        // no-op
    }
}
