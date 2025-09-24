#pragma once
// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.

#include <system/io/stream.h>
#include <system/array.h>
#include <cstdint>
#include <Aspose.Words.Cpp/Model/Document/Document.h>

namespace Aspose {

namespace Words {

namespace ApiExamples {

namespace TestData {

namespace TestClasses {

class DocumentTestClass : public System::Object
{
    typedef DocumentTestClass ThisType;
    typedef System::Object BaseType;
    
    typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
    RTTI_INFO_DECL();
    
public:

    System::SharedPtr<Aspose::Words::Document> get_Document() const;
    void set_Document(System::SharedPtr<Aspose::Words::Document> value);
    System::SharedPtr<System::IO::Stream> get_DocumentStream() const;
    void set_DocumentStream(System::SharedPtr<System::IO::Stream> value);
    System::ArrayPtr<uint8_t> get_DocumentBytes() const;
    void set_DocumentBytes(System::ArrayPtr<uint8_t> value);
    System::String get_DocumentString() const;
    void set_DocumentString(System::String value);
    
    DocumentTestClass(System::SharedPtr<Aspose::Words::Document> doc, System::SharedPtr<System::IO::Stream> docStream, System::ArrayPtr<uint8_t> docBytes, System::String docString);
    
private:

    System::SharedPtr<Aspose::Words::Document> pr_Document;
    System::SharedPtr<System::IO::Stream> pr_DocumentStream;
    System::ArrayPtr<uint8_t> pr_DocumentBytes;
    System::String pr_DocumentString;
    
};

} // namespace TestClasses
} // namespace TestData
} // namespace ApiExamples
} // namespace Words
} // namespace Aspose


