# C++ API for Word Document Processing

A standalone C++ class library to load, save, format & convert Microsoft Word® documents without Office Automation. [Aspose.Words for C++](https://products.aspose.com/words/cpp) is a powerful on premise API that can be used for a great range of document processing tasks such as; working with documents, sections, tables, bookmarks, fields, form fields, DocumentBuilder, ranges, comments, content control SDT, images, styles, charts, watermarks, lists, shapes, mail merge, rendering, printing and much more.

<p align="center">

  <a title="Download complete Aspose.Words for C++ source code" href="https://github.com/aspose-words/Aspose.Words-for-C/archive/master.zip">
	<img src="https://raw.github.com/AsposeExamples/java-examples-dashboard/master/images/downloadZip-Button-Large.png" />
  </a>
</p>

## Word File Processing via C++

- Detect the file format and [check file format compatibility](https://docs.aspose.com/words/cpp/creating-or-loading-a-document/).
- [Convert Word Documents](https://docs.aspose.com/words/cpp/converting-a-document/) to other supported formats.
- [Serialize a document object](https://docs.aspose.com/words/cpp/converting-a-document/) to fetch a byte array and vice versa.
- Supports conversion to various PDF standards.
- Export fonts to HTML in `Base64` encoding.
- [Save password encrypted Word document](https://docs.aspose.com/words/cpp/working-with-ooxml/#encrypt-document-with-password).
- Specify the OOXML specification.
- [Compress all the metafiles](https://docs.aspose.com/words/cpp/working-with-saveoptions/#compress-metafiles), be it small or large.
- Enable bi-directional marks to add support for right to left languages.
- Access the VBA Project to extend the functionality.
- Read, write or modify [VBA Macros](https://docs.aspose.com/words/cpp/working-with-vba-macros/).

## Read & Write Word Processing Files

**Microsoft Word:** DOC, DOT, DOCX, DOCM, DOTX, DOTM, RTF, FlatOPC, FlatOpcMacroEnabled, FlatOpcTemplate, FlatOpcTemplateMacroEnabled\
**WordprocessingML:** WordML\
**Web:** HTML, MHTML\
**OpenOffice:** ODT\
**Other:** TXT, MOBI

## Save Word Documents As

**Fixed Layout:** PDF, XPS, OpenXPS\
**Graphics:** SVG, EMF\
**Web:** HtmlFixed\
**Other:** PS, PCL, XamlFlow, XamlFixed, EPUB

## Get Started with Aspose.Words for C++

Are you ready to give Aspose.Words for C++ a try? Simply execute `Install-Package Aspose.Words.Cpp` from Package Manager Console in Visual Studio to fetch the NuGet package. If you already have Aspose.Words for C++ and want to upgrade the version, please execute `Update-Package Aspose.Words.Cpp` to get the latest version.

## Create DOCX with HelloWorld Text via C++

```c++
// create a blank document.
System::SharedPtr<Document> doc = System::MakeObject<Document>();
// the DocumentBuilder class provides members to easily add content to a document.
System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);
// write a new paragraph in the document with the text "Hello World!"
builder->Writeln(u"Hello World!");
// save the document. 
// the format to save as is inferred from the extension of the file name.
doc->Save(u"output.docx");
```

## Convert DOC to EPUB Format

```c++
// load the document from disk.
System::SharedPtr<Document> doc = System::MakeObject<Document>(u"template.doc");
// save the document in EPUB format.
doc->Save(u"output.epub");
```

[Home](https://www.aspose.com/) | [Product Page](https://products.aspose.com/words/cpp) | [Docs](https://docs.aspose.com/words/cpp/) | [Demos](https://products.aspose.app/words/family) | [API Reference](https://apireference.aspose.com/words/cpp) | [Examples](https://github.com/aspose-words/Aspose.Words-for-C) | [Blog](https://blog.aspose.com/category/words/) | [Free Support](https://forum.aspose.com/c/words) |  [Temporary License](https://purchase.aspose.com/temporary-license)
