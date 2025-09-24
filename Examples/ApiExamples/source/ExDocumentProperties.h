#pragma once
// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <system/object.h>
#include <cstdint>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Layout/Public/LayoutEnumerator.h>

#include "ApiExampleBase.h"


using namespace Aspose::Words::Layout;

namespace Aspose {

namespace Words {

namespace ApiExamples {

class ExDocumentProperties : public ApiExampleBase
{
    typedef ExDocumentProperties ThisType;
    typedef ApiExampleBase BaseType;
    
    typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
    RTTI_INFO_DECL();
    
private:

    /// <summary>
    /// Counts the lines in a document.
    /// Traverses the document's layout entities tree upon construction,
    /// counting entities of the "Line" type that also contain real text.
    /// </summary>
    class LineCounter : public System::Object
    {
        typedef LineCounter ThisType;
        typedef System::Object BaseType;
        
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();
        
    public:
    
        LineCounter(System::SharedPtr<Aspose::Words::Document> doc);
        
        int32_t GetLineCount();
        
    private:
    
        System::SharedPtr<Aspose::Words::Layout::LayoutEnumerator> mLayoutEnumerator;
        int32_t mLineCount;
        bool mScanningLineForRealText;
        
        void CountLines();
        
    };
    
    
public:

    void BuiltIn();
    void Custom();
    void Description();
    void Origin();
    void Content();
    void Thumbnail();
    void HyperlinkBase();
    void HeadingPairs();
    void Security();
    void CustomNamedAccess();
    void LinkCustomDocumentPropertiesToBookmark();
    void DocumentPropertyCollection();
    void PropertyTypes();
    void ExtendedProperties();
    
protected:

    void TestContent(System::SharedPtr<Aspose::Words::Document> doc);
    
};

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose


