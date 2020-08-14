#pragma once
// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <cstdint>

#include "ApiExampleBase.h"

namespace Aspose { namespace Words { namespace Layout { class LayoutEnumerator; } } }
namespace Aspose { namespace Words { class Document; } }

namespace ApiExamples {

class ExProperties : public ApiExampleBase
{
private:

    /// <summary>
    /// Util class that counts the lines in a document.
    /// Upon construction, traverses the document's layout entities tree,
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
        
    protected:
    
        System::Object::shared_members_type GetSharedMembers() override;
        
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
    
protected:

    void TestContent(System::SharedPtr<Aspose::Words::Document> doc);
    
};

} // namespace ApiExamples


