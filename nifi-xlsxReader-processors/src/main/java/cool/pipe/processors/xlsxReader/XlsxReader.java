/*
 * Licensed to the Apache Software Foundation (ASF) under one or more
 * contributor license agreements.  See the NOTICE file distributed with
 * this work for additional information regarding copyright ownership.
 * The ASF licenses this file to You under the Apache License, Version 2.0
 * (the "License"); you may not use this file except in compliance with
 * the License.  You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */
package cool.pipe.processors.xlsxReader;

import cool.pipe.processors.utils.UtilsKt;
import org.apache.nifi.components.PropertyDescriptor;
import org.apache.nifi.flowfile.FlowFile;
import org.apache.nifi.annotation.behavior.ReadsAttribute;
import org.apache.nifi.annotation.behavior.ReadsAttributes;
import org.apache.nifi.annotation.behavior.WritesAttribute;
import org.apache.nifi.annotation.behavior.WritesAttributes;
import org.apache.nifi.annotation.documentation.CapabilityDescription;
import org.apache.nifi.annotation.documentation.SeeAlso;
import org.apache.nifi.annotation.documentation.Tags;
import org.apache.nifi.processor.AbstractProcessor;
import org.apache.nifi.processor.ProcessContext;
import org.apache.nifi.processor.ProcessSession;
import org.apache.nifi.processor.ProcessorInitializationContext;
import org.apache.nifi.processor.Relationship;
import org.apache.nifi.processor.util.StandardValidators;

import java.util.ArrayList;
import java.util.Collections;
import java.util.HashSet;
import java.util.List;
import java.util.Set;

@Tags({"xlsx", "MSExcel", "Excel", "JSON"})
@CapabilityDescription("Reads a Microsoft Excel file and transforms it into JSON.")
@SeeAlso()
@ReadsAttributes({@ReadsAttribute(attribute = "")})
@WritesAttributes({@WritesAttribute(attribute = "")})
public class XlsxReader extends AbstractProcessor {

    public static final PropertyDescriptor PATH = new PropertyDescriptor
            .Builder().name("path")
            .displayName("Path")
            .description("The path to the Excel file to be used, e.g., '/home/username/myfile.xlsx'.")
            .required(true)
            .addValidator(StandardValidators.FILE_EXISTS_VALIDATOR)
            .build();

    public static final PropertyDescriptor RANGE = new PropertyDescriptor
            .Builder().name("range")
            .displayName("Range")
            .description("A1:B1")
            .required(true)
            .addValidator(StandardValidators.NON_EMPTY_VALIDATOR)
            .build();

    public static final PropertyDescriptor SHEET_INDEX = new PropertyDescriptor
            .Builder().name("sheet_index")
            .displayName("Sheet Index")
            .description("0")
            .required(true)
            .defaultValue("0")
            .addValidator(StandardValidators.NUMBER_VALIDATOR)
            .build();

    public static final PropertyDescriptor HEADERS = new PropertyDescriptor
            .Builder().name("headers")
            .displayName("Headers")
            .description("Accepts true or false.")
            .required(true)
            .allowableValues("true", "false")
            .defaultValue("true")
            .build();

    public static final PropertyDescriptor FORMAT_DATE_HEADER = new PropertyDescriptor
            .Builder().name("FormatoFechaHeader")
            .displayName("Header Date Format")
            .description("This is the date format to be used if there is a date in the header. By default, it will have the value 'yyyy-MM-dd'.")
            .required(true)
            .defaultValue("yyyy-MM-dd")
            .addValidator(StandardValidators.NON_EMPTY_VALIDATOR)
            .build();

    public static final PropertyDescriptor FORMAT_DATE_BODY = new PropertyDescriptor
            .Builder().name("FormatoFechaBody")
            .displayName("Body Date Format")
            .description("This is the date format to be used if there is a date in the body. By default, it will have the value 'yyyy-MM-dd HH:mm:ss.SSS'.")
            .required(true)
            .defaultValue("yyyy-MM-dd HH:mm:ss.SSS")
            .addValidator(StandardValidators.NON_EMPTY_VALIDATOR)
            .build();

    public static final Relationship REL_SUCCESS = new Relationship.Builder()
            .name("REL_SUCCESS")
            .description("A Flowfile is routed to this relationship when everything goes well here.")
            .build();

    public static final Relationship REL_FAILURE = new Relationship.Builder()
            .name("REL_FAILURE")
            .description("A Flowfile is routed to this relationship when something is invalid.")
            .build();

    private List<PropertyDescriptor> descriptors;

    private Set<Relationship> relationships;

    @Override
    protected void init(final ProcessorInitializationContext context) {
        descriptors = new ArrayList<>();
        descriptors.add(PATH);
        descriptors.add(RANGE);
        descriptors.add(SHEET_INDEX);
        descriptors.add(HEADERS);
        descriptors.add(FORMAT_DATE_HEADER);
        descriptors.add(FORMAT_DATE_BODY);
        descriptors = Collections.unmodifiableList(descriptors);

        relationships = new HashSet<>();
        relationships.add(REL_SUCCESS);
        relationships.add(REL_FAILURE);
        relationships = Collections.unmodifiableSet(relationships);
    }

    @Override
    public Set<Relationship> getRelationships() {
        return this.relationships;
    }

    @Override
    public final List<PropertyDescriptor> getSupportedPropertyDescriptors() {
        return descriptors;
    }

    @Override
    public void onTrigger(final ProcessContext context, final ProcessSession session) {
        FlowFile flowFile = session.get();
        if (flowFile == null) {
            flowFile = session.create();
        }

        //leemos las propiedades
        String path = context.getProperty(PATH).evaluateAttributeExpressions(flowFile).getValue();
        String range = context.getProperty(RANGE).evaluateAttributeExpressions(flowFile).getValue();
        int sheet_index = Integer.parseInt(context.getProperty(SHEET_INDEX).evaluateAttributeExpressions(flowFile).getValue());
        boolean headers = Boolean.parseBoolean(context.getProperty(HEADERS).evaluateAttributeExpressions(flowFile).getValue());
        String formatHeader = context.getProperty(FORMAT_DATE_HEADER).evaluateAttributeExpressions(flowFile).getValue();
        String formatBody = context.getProperty(FORMAT_DATE_BODY).evaluateAttributeExpressions(flowFile).getValue();


        //obtenemos el resultado
        String result = UtilsKt.readXlsx(path, range, sheet_index, headers, formatHeader, formatBody);

        //revisamos si el string es vacío o no
        if (result.isBlank()) {
            flowFile = session.putAttribute(flowFile, "status", String.valueOf(false));
            getLogger().error("No se encontraron datos para leer.");
            session.transfer(flowFile, REL_FAILURE);
        } else {
            flowFile = session.putAttribute(flowFile, "mime.type", "application/json");

            flowFile = session.write(flowFile, outputStream -> outputStream.write(result.getBytes()));
            session.transfer(flowFile, REL_SUCCESS);
        }

        session.commit();
    }
}
