# NiFi Custom Processor - nifi-xlsxReader-bundle
[![en](https://img.shields.io/badge/lang-en-red.svg)](https://github.com/jonatasemidio/multilanguage-readme-pattern/blob/master/README.md)

This custom processor for Apache NiFi, called `nifi-xlsxReader-bundle`, has been developed to allow reading Excel files in XLSX format and converting the read content into a `FlowFile` that can be used in the NiFi workflow.

## Processor Details

- **Processor Name**: nifi-xlsxReader-bundle
- **Programming Language**: Kotlin
- Library Used**: Apache POI
- **Compile Command**: `mvn clean install`: `mvn clean install`.

## Functionality

The `nifi-xlsxReader-bundle` processor allows to read Excel XLSX files and convert their content into a `FlowFile`. This is especially useful in NiFi workflows where you need to process data contained in Excel spreadsheets.

## Library Used

The main library used to perform the reading of Excel XLSX files is [Apache POI](https://poi.apache.org/), a powerful Java library that provides classes and methods for working with Microsoft Office file formats, including Excel files.

## Compiling the Processor

To compile the custom processor `nifi-xlsxReader-bundle`, the following command must be executed in the project directory:

```
mvn clean install
```
This command will compile the source code, generate the processor JAR file and prepare it for deployment in Apache NiFi.

With nifi-xlsxReader-bundle, Apache NiFi users can easily integrate the reading of Excel XLSX files into their data workflows. This provides greater flexibility and versatility in real-time data manipulation.

If you would like to learn more about this custom processor or would like to contribute to the project, please visit the project's GitHub repository.

Enjoy the power of Apache NiFi with the nifi-xlsxReader-bundle processor!