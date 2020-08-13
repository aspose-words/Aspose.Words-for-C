#include "ExSmartTag.h"

#include <testing/test_predicates.h>
#include <system/test_tools/test_tools.h>
#include <system/test_tools/compare.h>
#include <system/string.h>
#include <system/shared_ptr.h>
#include <system/object.h>
#include <system/exceptions.h>
#include <system/enumerator_adapter.h>
#include <system/details/dispose_guard.h>
#include <system/console.h>
#include <system/collections/ienumerator.h>
#include <system/array.h>
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
#include <Aspose.Words.Cpp/Model/Markup/SmartTag.h>
#include <Aspose.Words.Cpp/Model/Markup/CustomXmlPropertyCollection.h>
#include <Aspose.Words.Cpp/Model/Markup/CustomXmlProperty.h>
#include <Aspose.Words.Cpp/Model/Document/VisitorAction.h>
#include <Aspose.Words.Cpp/Model/Document/SaveFormat.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>

using System::String;
using System::SharedPtr;
using System::ArrayPtr;
using System::MakeObject;
using System::MakeArray;

using namespace Aspose::Words;
using namespace Aspose::Words::Markup;
namespace ApiExamples {

RTTI_INFO_IMPL_HASH(1293543254u, ::ApiExamples::ExSmartTag::SmartTagVisitor, ThisTypeBaseTypesInfo);

Aspose::Words::VisitorAction ExSmartTag::SmartTagVisitor::VisitSmartTagStart(SharedPtr<Aspose::Words::Markup::SmartTag> smartTag)
{
    System::Console::WriteLine(String::Format(u"Smart tag type: {0}",smartTag->get_Element()));
    return Aspose::Words::VisitorAction::Continue;
}

Aspose::Words::VisitorAction ExSmartTag::SmartTagVisitor::VisitSmartTagEnd(SharedPtr<Aspose::Words::Markup::SmartTag> smartTag)
{
    System::Console::WriteLine(String::Format(u"\tContents: \"{0}\"",smartTag->ToString(Aspose::Words::SaveFormat::Text)));

    if (smartTag->get_Properties()->get_Count() == 0)
    {
        System::Console::WriteLine(u"\tContains no properties");
    }
    else
    {
        System::Console::Write(u"\tProperties: ");
        auto properties = MakeArray<String>(smartTag->get_Properties()->get_Count());
        int index = 0;

        for (auto cxp : System::IterateOver(smartTag->get_Properties()))
        {
            properties[index++] = String::Format(u"\"{0}\" = \"{1}\"",cxp->get_Name(),cxp->get_Value());
        }

        System::Console::WriteLine(String::Join(u", ", properties));
    }

    return Aspose::Words::VisitorAction::Continue;
}

void ExSmartTag::TestCreate(SharedPtr<Aspose::Words::Document> doc)
{
    auto smartTag = System::DynamicCast<Aspose::Words::Markup::SmartTag>(doc->GetChild(Aspose::Words::NodeType::SmartTag, 0, true));

    ASSERT_EQ(u"date", smartTag->get_Element());
    ASSERT_EQ(u"May 29, 2019", smartTag->GetText());
    ASSERT_EQ(u"urn:schemas-microsoft-com:office:smarttags", smartTag->get_Uri());

    ASSERT_EQ(u"Day", smartTag->get_Properties()->idx_get(0)->get_Name());
    ASSERT_EQ(String::Empty, smartTag->get_Properties()->idx_get(0)->get_Uri());
    ASSERT_EQ(u"29", smartTag->get_Properties()->idx_get(0)->get_Value());
    ASSERT_EQ(u"Month", smartTag->get_Properties()->idx_get(1)->get_Name());
    ASSERT_EQ(String::Empty, smartTag->get_Properties()->idx_get(1)->get_Uri());
    ASSERT_EQ(u"5", smartTag->get_Properties()->idx_get(1)->get_Value());
    ASSERT_EQ(u"Year", smartTag->get_Properties()->idx_get(2)->get_Name());
    ASSERT_EQ(String::Empty, smartTag->get_Properties()->idx_get(2)->get_Uri());
    ASSERT_EQ(u"2019", smartTag->get_Properties()->idx_get(2)->get_Value());

    smartTag = System::DynamicCast<Aspose::Words::Markup::SmartTag>(doc->GetChild(Aspose::Words::NodeType::SmartTag, 1, true));

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
    static SharedPtr<::ApiExamples::ExSmartTag> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

    static void SetUpTestCase()
    {
        s_instance = MakeObject<::ApiExamples::ExSmartTag>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };

};

SharedPtr<::ApiExamples::ExSmartTag> ExSmartTag::s_instance;

} // namespace gtest_test

void ExSmartTag::Create()
{
    auto doc = MakeObject<Document>();
    auto smartTag = MakeObject<SmartTag>(doc);
    smartTag->set_Element(u"date");

    // Specify a date and set smart tag properties accordingly
    smartTag->AppendChild(MakeObject<Run>(doc, u"May 29, 2019"));

    smartTag->get_Properties()->Add(MakeObject<CustomXmlProperty>(u"Day", String::Empty, u"29"));
    smartTag->get_Properties()->Add(MakeObject<CustomXmlProperty>(u"Month", String::Empty, u"5"));
    smartTag->get_Properties()->Add(MakeObject<CustomXmlProperty>(u"Year", String::Empty, u"2019"));

    // Set the smart tag's uri to the default
    smartTag->set_Uri(u"urn:schemas-microsoft-com:office:smarttags");

    doc->get_FirstSection()->get_Body()->get_FirstParagraph()->AppendChild(smartTag);
    doc->get_FirstSection()->get_Body()->get_FirstParagraph()->AppendChild(MakeObject<Run>(doc, u" is a date. "));

    // Create and add one more smart tag, this time for a financial symbol
    smartTag = MakeObject<SmartTag>(doc);
    smartTag->set_Element(u"stockticker");
    smartTag->set_Uri(u"urn:schemas-microsoft-com:office:smarttags");

    smartTag->AppendChild(MakeObject<Run>(doc, u"MSFT"));

    doc->get_FirstSection()->get_Body()->get_FirstParagraph()->AppendChild(smartTag);
    doc->get_FirstSection()->get_Body()->get_FirstParagraph()->AppendChild(MakeObject<Run>(doc, u" is a stock ticker."));

    // Print all the smart tags in our document with a document visitor
    doc->Accept(MakeObject<ExSmartTag::SmartTagVisitor>());

    // SmartTags are supported by older versions of microsoft Word
    doc->Save(ArtifactsDir + u"SmartTag.Create.doc");

    // We can strip a document of all its smart tags with RemoveSmartTags()
    ASSERT_EQ(2, doc->GetChildNodes(Aspose::Words::NodeType::SmartTag, true)->get_Count());
    doc->RemoveSmartTags();
    ASSERT_EQ(0, doc->GetChildNodes(Aspose::Words::NodeType::SmartTag, true)->get_Count());

    TestCreate(MakeObject<Document>(ArtifactsDir + u"SmartTag.Create.doc"));
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
    // Open a document that contains smart tags and their collection
    auto doc = MakeObject<Document>(MyDir + u"Smart tags.doc");

    // Smart tags are an older Microsoft Word feature that can automatically detect and tag
    // any parts of the text that it registers as commonly used information objects such as names, addresses, stock tickers, dates etc
    // In Word 2003, smart tags can be turned on in Tools > AutoCorrect options... > SmartTags tab
    // In our input document there are three objects that were registered as smart tags, but since they can be nested, we have 8 in this collection
    SharedPtr<NodeCollection> smartTags = doc->GetChildNodes(Aspose::Words::NodeType::SmartTag, true);
    ASSERT_EQ(8, smartTags->get_Count());

    // The last smart tag is of the "Date" type, which we will retrieve here
    auto smartTag = System::DynamicCast<Aspose::Words::Markup::SmartTag>(smartTags->idx_get(7));

    // The Properties attribute, for some smart tags, elaborates on the text object that Word picked up as a smart tag
    // In the case of our "Date" smart tag, its properties will let us know the year, month and day within the smart tag
    SharedPtr<CustomXmlPropertyCollection> properties = smartTag->get_Properties();

    // We can enumerate over the collection and print the aforementioned properties to the console
    ASSERT_EQ(4, properties->get_Count());

    {
        SharedPtr<System::Collections::Generic::IEnumerator<SharedPtr<CustomXmlProperty>>> enumerator = properties->GetEnumerator();
        while (enumerator->MoveNext())
        {
            System::Console::WriteLine(String::Format(u"Property name: {0}, value: {1}",enumerator->get_Current()->get_Name(),enumerator->get_Current()->get_Value()));
            ASSERT_EQ(u"", enumerator->get_Current()->get_Uri());
        }
    }

    // We can also access the elements in various ways, including as a key-value pair
    ASSERT_TRUE(properties->Contains(u"Day"));
    ASSERT_EQ(u"22", properties->idx_get(u"Day")->get_Value());
    ASSERT_EQ(u"2003", properties->idx_get(2)->get_Value());
    ASSERT_EQ(1, properties->IndexOfKey(u"Month"));

    // We can also remove elements by name, index or clear the collection entirely
    properties->RemoveAt(3);
    properties->Remove(u"Year");
    ASSERT_EQ(2, (properties->get_Count()));

    properties->Clear();
    ASSERT_EQ(0, (properties->get_Count()));

    // We can remove the entire smart tag like this
    smartTag->Remove();
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
