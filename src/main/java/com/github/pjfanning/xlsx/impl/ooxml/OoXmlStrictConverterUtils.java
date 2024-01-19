package com.github.pjfanning.xlsx.impl.ooxml;

import com.github.pjfanning.xlsx.exceptions.ReadException;
import org.apache.poi.util.Internal;

import java.io.*;
import java.nio.charset.StandardCharsets;
import java.util.Properties;

@Internal
final class OoXmlStrictConverterUtils {

    private OoXmlStrictConverterUtils() {}

    static boolean isBlank(final String str) {
        return str == null || str.trim().length() == 0;
    }

    static boolean isNotBlank(final String str) {
        return !isBlank(str);
    }

    static Properties readMappings() throws ReadException {
        Properties props = new Properties();
        try(InputStream is = OoXmlStrictConverterUtils.class.getResourceAsStream("/ooxml-strict-mappings.properties");
                BufferedReader reader = new BufferedReader(new InputStreamReader(is, StandardCharsets.ISO_8859_1))) {
            String line;
            while((line = reader.readLine()) != null) {
                String[] vals = line.split("=");
                if(vals.length >= 2) {
                    props.setProperty(vals[0], vals[1]);
                } else if(vals.length == 1) {
                    props.setProperty(vals[0], "");
                }

            }
        } catch (IOException e) {
            throw new ReadException("Failed to read mappings", e);
        }
        return props;
    }

}
