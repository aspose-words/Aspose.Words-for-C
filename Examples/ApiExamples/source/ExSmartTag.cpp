// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExSmartTag.h"

#include <testing/test_predicates.h>
#include <system/test_tools/test_tools.h>
#include <system/test_tools/compare.h>
#include <system/object_ext.h>
#include <system/linq/enumerable.h>
#include <system/enumerator_adapter.h>
#include <system/details/dispose_guard.h>
#include <system/collections/ienumerator.h>
#include <system/collections/ienumerable.h>
#include <system/array.h>
#include <iostream>
#include <gtest/gtest.h>
#include <cstdint>
#include <Aspose.Words.Cpp/Model/Text/Run.h>
#include <Aspose.Words.Cpp/Model/Text/Paragraph.h>
#include <Aspose.Words.Cpp/Model/Sections/Section.h>
#include <Aspose.Words.Cpp/Model/Sections/Body.h>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeType.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeCollection.h>
#include <Aspose.Words.Cpp/Model/Nodes/Node.h>
#include <Aspose.Words.Cpp/Model/Markup/CustomXmlPropertyCollection.h>
#include <Aspose.Words.Cpp/Model/Markup/CustomXmlProperty.h>
#include <Aspose.Words.Cpp/Model/Document/SaveFormat.h>


using namespace Aspose::Words::Markup;
namespace Aspose {

namespace Words {

namespace ApiExamples {

RTTI_INFO_IMPL_HASH(3831900680u, ::Aspose::Words::ApiExamples::ExSmartTag::SmartTagPrinter, ThisTypeBaseTypesInfo);

Aspose::Words::VisitorAction ExSmartTag::SmartTagPrinter::VisitSmartTagStart(System::SharedPtr<Aspose::Words::Markup::SmartTag> smartTag)
{
    std::cout << System::String::Format(u"Smart tag type: {0}", smartTag->get_Element()) << std::endl;
    return Aspose::Words::VisitorAction::Continue;
}

Aspose::Words::VisitorAction ExSmartTag::SmartTagPrinter::VisitSmartTagEnd(System::SharedPtr<Aspose::Words::Markup::SmartTag> smartTag)
{
    std::cout << System::String::Format(u"\tContents: \"{0}\"", smartTag->ToString(Aspose::Words::SaveFormat::Text)) << std::endl;
    
    if (smartTag->get_Properties()->get_Count() == 0)
    {
        std::cout << "\tContains no properties" << std::endl;
    }
    else
    {
        std::cout << "\tProperties: ";
        auto properties = System::MakeArray<System::String>(smartTag->get_Properties()->get_Count());
        int32_t index = 0;
        
        for (auto&& cxp : System::IterateOver(smartTag->get_Properties()))
        {
            properties[index++] = System::String::Format(u"\"{0}\" = \"{1}\"", cxp->get_Name(), cxp->get_Value());
        }
        
        std::cout << System::String::Join(u", ", properties) << std::endl;
    }
    
    return Aspose::Words::VisitorAction::Continue;
}


RTTI_INFO_IMPL_HASH(535207617u, ::Aspose::Words::ApiExamples::ExSmartTag, ThisTypeBaseTypesInfo);

void ExSmartTag::TestCreate(System::SharedPtr<Aspose::Words::Document> doc)
{
    auto smartTag = System::ExplicitCast<Aspose::Words::Markup::SmartTag>(doc->GetChild(Aspose::Words::NodeType::SmartTag, 0, true));
    
    ASSERT_EQ(u"date", smartTag->get_Element());
    ASSERT_EQ(u"May 29, 2019", smartTag->GetText());
    ASSERT_EQ(u"urn:schemas-microsoft-com:office:smarttags", smartTag->get_Uri());
    
    ASSERT_EQ(u"Day", smartTag->get_Properties()->idx_get(0)->get_Name());
    ASSERT_EQ(System::String::Empty, smartTag->get_Properties()->idx_get(0)->get_Uri());
    ASSERT_EQ(u"29", smartTag->get_Properties()->idx_get(0)->get_Value());
    ASSERT_EQ(u"Month", smartTag->get_Properties()->idx_get(1)->get_Name());
    ASSERT_EQ(System::String::Empty, smartTag->get_Properties()->idx_get(1)->get_Uri());
    ASSERT_EQ(u"5", smartTag->get_Properties()->idx_get(1)->get_Value());
    ASSERT_EQ(u"Year", smartTag->get_Properties()->idx_get(2)->get_Name());
    ASSERT_EQ(System::String::Empty, smartTag->get_Properties()->idx_get(2)->get_Uri());
    ASSERT_EQ(u"2019", smartTag->get_Properties()->idx_get(2)->get_Value());
    
    smartTag = System::ExplicitCast<Aspose::Words::Markup::SmartTag>(doc->GetChild(Aspose::Words::NodeType::SmartTag, 1, true));
    
    ASSERT_EQ(u"stockticker", smartTag->get_Element());
    ASSERT_EQ(u"MSFT", smartTag->GetText());
    ASSERT_EQ(u"urn:schemas-microsoft-com:office:smarttags", smartTag->get_Uri());
    ASSERT_EQ(0, smartTag->get_Properties()->get_Count());
}


namespace gtest_test
{

class ExSmartTag : public ::testing::Test
{
protected:
    static System::SharedPtr<::Aspose::Words::ApiExamples::ExSmartTag> s_instance;
    
    void SetUp() override
    {
        s_instance->SetUp();
    };
    
    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::Aspose::Words::ApiExamples::ExSmartTag>();
        s_instance->OneTimeSetUp();
    };
    
    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
    
};

System::SharedPtr<::Aspose::Words::ApiExamples::ExSmartTag> ExSmartTag::s_instance;

} // namespace gtest_test

void ExSmartTag::Create()
{
    auto doc = System::MakeObject<Aspose::Words::Document>();
    
    // A smart tag appears in a document with Microsoft Word recognizes a part of its text as some form of data,
    // such as a name, date, or address, and converts it to a hyperlink that displays a purple dotted underline.
    auto smartTag = System::MakeObject<Aspose::Words::Markup::SmartTag>(doc);
    
    // Smart tags are composite nodes that contain their recognized text in its entirety.
    // Add contents to this smart tag manually.
    smartTag->AppendChild<System::SharedPtr<Aspose::Words::Run>>(System::MakeObject<Aspose::Words::Run>(doc, u"May 29, 2019"));
    
    // Microsoft Word may recognize the above contents as being a date.
    // Smart tags use the "Element" property to reflect the type of data they contain.
    smartTag->set_Element(u"date");
    
    // Some smart tag types process their contents further into custom XML properties.
    smartTag->get_Properties()->Add(System::MakeObject<Aspose::Words::Markup::CustomXmlProperty>(u"Day", System::String::Empty, u"29"));
    smartTag->get_Properties()->Add(System::MakeObject<Aspose::Words::Markup::CustomXmlProperty>(u"Month", System::String::Empty, u"5"));
    smartTag->get_Properties()->Add(System::MakeObject<Aspose::Words::Markup::CustomXmlProperty>(u"Year", System::String::Empty, u"2019"));
    
    // Set the smart tag's URI to the default value.
    smartTag->set_Uri(u"urn:schemas-microsoft-com:office:smarttags");
    
    doc->get_FirstSection()->get_Body()->get_FirstParagraph()->AppendChild<System::SharedPtr<Aspose::Words::Markup::SmartTag>>(smartTag);
    doc->get_FirstSection()->get_Body()->get_FirstParagraph()->AppendChild<System::SharedPtr<Aspose::Words::Run>>(System::MakeObject<Aspose::Words::Run>(doc, u" is a date. "));
    
    // Create another smart tag for a stock ticker.
    smartTag = System::MakeObject<Aspose::Words::Markup::SmartTag>(doc);
    smartTag->set_Element(u"stockticker");
    smartTag->set_Uri(u"urn:schemas-microsoft-com:office:smarttags");
    
    smartTag->AppendChild<System::SharedPtr<Aspose::Words::Run>>(System::MakeObject<Aspose::Words::Run>(doc, u"MSFT"));
    
    doc->get_FirstSection()->get_Body()->get_FirstParagraph()->AppendChild<System::SharedPtr<Aspose::Words::Markup::SmartTag>>(smartTag);
    doc->get_FirstSection()->get_Body()->get_FirstParagraph()->AppendChild<System::SharedPtr<Aspose::Words::Run>>(System::MakeObject<Aspose::Words::Run>(doc, u" is a stock ticker."));
    
    // Print all the smart tags in our document using a document visitor.
    doc->Accept(System::MakeObject<Aspose::Words::ApiExamples::ExSmartTag::SmartTagPrinter>());
    
    // Older versions of Microsoft Word support smart tags.
    doc->Save(get_ArtifactsDir() + u"SmartTag.Create.doc");
    
    // Use the "RemoveSmartTags" method to remove all smart tags from a document.
    ASSERT_EQ(2, doc->GetChildNodes(Aspose::Words::NodeType::SmartTag, true)->get_Count());
    
    doc->RemoveSmartTags();
    
    ASSERT_EQ(0, doc->GetChildNodes(Aspose::Words::NodeType::SmartTag, true)->get_Count());
    TestCreate(System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"SmartTag.Create.doc"));
    //ExSkip
}

namespace gtest_test
{

TEST_F(ExSmartTag, Create)
{
    s_instance->Create();
}

} // namespace gtest_test

void ExSmartTag::Properties()
{
    //ExStart
    //ExFor:CustomXmlProperty.Uri
    //ExFor:CustomXmlPropertyCollection
    //ExFor:CustomXmlPropertyCollection.Add(CustomXmlProperty)
    //ExFor:CustomXmlPropertyCollection.Clear
    //ExFor:CustomXmlPropertyCollection.Contains(String)
    //ExFor:CustomXmlPropertyCollection.Count
    //ExFor:CustomXmlPropertyCollection.GetEnumerator
    //ExFor:CustomXmlPropertyCollection.IndexOfKey(String)
    //ExFor:CustomXmlPropertyCollection.Item(Int32)
    //ExFor:CustomXmlPropertyCollection.Item(String)
    //ExFor:CustomXmlPropertyCollection.Remove(String)
    //ExFor:CustomXmlPropertyCollection.RemoveAt(Int32)
    //ExSummary:Shows how to work with smart tag properties to get in depth information about smart tags.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Smart tags.doc");
    
    // A smart tag appears in a document with Microsoft Word recognizes a part of its text as some form of data,
    // such as a name, date, or address, and converts it to a hyperlink that displays a purple dotted underline.
    // In Word 2003, we can enable smart tags via "Tools" -> "AutoCorrect options..." -> "SmartTags".
    // In our input document, there are three objects that Microsoft Word registered as smart tags.
    // Smart tags may be nested, so this collection contains more.
    System::ArrayPtr<System::SharedPtr<Aspose::Words::Markup::SmartTag>> smartTags = doc->GetChildNodes(Aspose::Words::NodeType::SmartTag, true)->LINQ_OfType<System::SharedPtr<Aspose::Words::Markup::SmartTag> >()->LINQ_ToArray();
    
    ASSERT_EQ(8, smartTags->get_Length());
    
    // The "Properties" member of a smart tag contains its metadata, which will be different for each type of smart tag.
    // The properties of a "date"-type smart tag contain its year, month, and day.
    System::SharedPtr<Aspose::Words::Markup::CustomXmlPropertyCollection> properties = smartTags[7]->get_Properties();
    
    ASSERT_EQ(4, properties->get_Count());
    
    {
        System::SharedPtr<System::Collections::Generic::IEnumerator<System::SharedPtr<Aspose::Words::Markup::CustomXmlProperty>>> enumerator = properties->GetEnumerator();
        while (enumerator->MoveNext())
        {
            std::cout << System::String::Format(u"Property name: {0}, value: {1}", enumerator->get_Current()->get_Name(), enumerator->get_Current()->get_Value()) << std::endl;
            ASSERT_EQ(u"", enumerator->get_Current()->get_Uri());
        }
    }
    
    // We can also access the properties in various ways, such as a key-value pair.
    ASSERT_TRUE(properties->Contains(u"Day"));
    ASSERT_EQ(u"22", properties->idx_get(u"Day")->get_Value());
    ASSERT_EQ(u"2003", properties->idx_get(2)->get_Value());
    ASSERT_EQ(1, properties->IndexOfKey(u"Month"));
    
    // Below are three ways of removing elements from the properties collection.
    // 1 -  Remove by index:
    properties->RemoveAt(3);
    
    ASSERT_EQ(3, properties->get_Count());
    
    // 2 -  Remove by name:
    properties->Remove(u"Year");
    
    ASSERT_EQ(2, properties->get_Count());
    
    // 3 -  Clear the entire collection at once:
    properties->Clear();
    
    ASSERT_EQ(0, properties->get_Count());
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExSmartTag, Properties)
{
    s_instance->Properties();
}

} // namespace gtest_test

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose
