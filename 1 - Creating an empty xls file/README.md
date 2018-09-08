# 1 - Creating an empty .xls file

The project shows how to create a simple .xls file (demo.xls). The file is empty and has a single sheet with the name Tab1.

## Dependencies

If you are completely sure that you need the .xls file type as final Excel file and you are not interested in .xlsx file you can use dependency with an artifact id poi. This dependency consists of all tools for .xls file format. It is, actually, a pure core poi.

```
<dependency>
  <groupId>org.apache.poi</groupId>
  <artifactId>poi</artifactId>
  <version>3.17</version>
</dependency>
```
## Result

The project has only one aim - showing how to create an empty .xls file with the Apache poi.

