# Procesador Personalizado NiFi - nifi-xlsxReader-bundle
[![en](https://img.shields.io/badge/lang-en-red.svg)](https://github.com/pipe-cool/nifi-xlsxReader-bundle/blob/main/README.en.md)

Este procesador personalizado para Apache NiFi, llamado `nifi-xlsxReader-bundle`, se ha desarrollado para permitir la lectura de archivos Excel en formato XLSX y convertir el contenido leído en un `JSON` que se puede utilizar en el flujo de trabajo de NiFi.

## Detalles del Procesador

- **Nombre del Procesador**: nifi-xlsxReader-bundle
- **Lenguaje de Programación**: Kotlin
- **Biblioteca Utilizada**: Apache POI
- **Comando de Compilación**: `mvn clean install`

## Funcionalidad

El procesador `nifi-xlsxReader-bundle` permite leer archivos Excel XLSX y convertir su contenido en un `JSON`. Esto es especialmente útil en flujos de trabajo de NiFi donde se necesita procesar datos contenidos en hojas de cálculo de Excel.

## Biblioteca Utilizada

La biblioteca principal utilizada para realizar la lectura de archivos Excel XLSX es [Apache POI](https://poi.apache.org/), una poderosa biblioteca de Java que proporciona clases y métodos para trabajar con formatos de archivo de Microsoft Office, incluidos archivos Excel.

## Compilación del Procesador

Para compilar el procesador personalizado `nifi-xlsxReader-bundle`, se debe ejecutar el siguiente comando en el directorio del proyecto:

```
mvn clean install
```
Este comando compilará el código fuente, generará el archivo JAR del procesador y lo preparará para su implementación en Apache NiFi.

Con nifi-xlsxReader-bundle, los usuarios de Apache NiFi pueden integrar fácilmente la lectura de archivos Excel XLSX en sus flujos de trabajo de datos. Esto proporciona una mayor flexibilidad y versatilidad en la manipulación de datos en tiempo real.

Si desea obtener más información sobre este procesador personalizado o desea contribuir al proyecto, visite el repositorio de GitHub del proyecto.

¡Disfrute de la potencia de Apache NiFi con el procesador nifi-xlsxReader-bundle!