#pragma once

#include <cstdint>
#include <Aspose.Words.Cpp/Document.h>
#include <system/array.h>
#include <system/io/stream.h>

using namespace Aspose::Words;

namespace ApiExamples { namespace TestData { namespace TestClasses {

class DocumentTestClass : public System::Object
{
public:
    System::SharedPtr<Aspose::Words::Document> get_Document();
    void set_Document(System::SharedPtr<Aspose::Words::Document> value);
    System::SharedPtr<System::IO::Stream> get_DocumentStream();
    void set_DocumentStream(System::SharedPtr<System::IO::Stream> value);
    System::ArrayPtr<uint8_t> get_DocumentBytes();
    void set_DocumentBytes(System::ArrayPtr<uint8_t> value);
    System::String get_DocumentString();
    void set_DocumentString(System::String value);

    DocumentTestClass(System::SharedPtr<Aspose::Words::Document> doc, System::SharedPtr<System::IO::Stream> docStream, System::ArrayPtr<uint8_t> docBytes,
                      System::String docString);

private:
    System::SharedPtr<Aspose::Words::Document> pr_Document;
    System::SharedPtr<System::IO::Stream> pr_DocumentStream;
    System::ArrayPtr<uint8_t> pr_DocumentBytes;
    System::String pr_DocumentString;
};

}}} // namespace ApiExamples::TestData::TestClasses
