#pragma once

#include <cstdint>
#include <iostream>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentVisitor.h>
#include <Aspose.Words.Cpp/Model/Document/SaveFormat.h>
#include <Aspose.Words.Cpp/Model/Document/VisitorAction.h>
#include <Aspose.Words.Cpp/Model/Markup/CustomXmlProperty.h>
#include <Aspose.Words.Cpp/Model/Markup/CustomXmlPropertyCollection.h>
#include <Aspose.Words.Cpp/Model/Markup/SmartTag.h>
#include <Aspose.Words.Cpp/Model/Nodes/Node.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeCollection.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeType.h>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Model/Sections/Body.h>
#include <Aspose.Words.Cpp/Model/Sections/Section.h>
#include <Aspose.Words.Cpp/Model/Text/Paragraph.h>
#include <Aspose.Words.Cpp/Model/Text/Run.h>
#include <system/array.h>
#include <system/collections/ienumerator.h>
#include <system/details/dispose_guard.h>
#include <system/enumerator_adapter.h>
#include <system/exceptions.h>
#include <system/test_tools/compare.h>
#include <system/test_tools/test_tools.h>
#include <testing/test_predicates.h>

#include "ApiExampleBase.h"

using System::String;
using System::SharedPtr;
using System::ArrayPtr;
using System::MakeObject;
using System::MakeArray;

using namespace Aspose::Words;
using namespace Aspose::Words::Markup;

namespace ApiExamples {

class ExSmartTag : public ApiExampleBase
{
public:
    //ExStart
    //ExFor:CompositeNode.RemoveSmartTags
    //ExFor:CustomXmlProperty
    //ExFor:CustomXmlProperty.#ctor(String,String,String)
    //ExFor:CustomXmlProperty.Name
    //ExFor:CustomXmlProperty.Value
    //ExFor:Markup.SmartTag
    //ExFor:Markup.SmartTag.#ctor(DocumentBase)
    //ExFor:Markup.SmartTag.Accept(DocumentVisitor)
    //ExFor:Markup.SmartTag.Element
    //ExFor:Markup.SmartTag.Properties
    //ExFor:Markup.SmartTag.Uri
    //ExSummary:Shows how to create smart tags.
    void Create()
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
        ASSERT_EQ(2, doc->GetChildNodes(NodeType::SmartTag, true)->get_Count());
        doc->RemoveSmartTags();
        ASSERT_EQ(0, doc->GetChildNodes(NodeType::SmartTag, true)->get_Count());

        TestCreate(MakeObject<Document>(ArtifactsDir + u"SmartTag.Create.doc"));
        //ExSkip
    }

private:
    /// <summary>
    /// DocumentVisitor implementation that prints smart tags and their contents
    /// </summary>
    class SmartTagVisitor : public DocumentVisitor
    {
    public:
        /// <summary>
        /// Called when a SmartTag node is encountered in the document.
        /// </summary>
        VisitorAction VisitSmartTagStart(SharedPtr<SmartTag> smartTag) override
        {
            std::cout << "Smart tag type: " << smartTag->get_Element() << std::endl;
            return VisitorAction::Continue;
        }

        /// <summary>
        /// Called when the visiting of a SmartTag node is ended.
        /// </summary>
        VisitorAction VisitSmartTagEnd(SharedPtr<SmartTag> smartTag) override
        {
            std::cout << "\tContents: \"" << smartTag->ToString(SaveFormat::Text) << "\"" << std::endl;

            if (smartTag->get_Properties()->get_Count() == 0)
            {
                std::cout << "\tContains no properties" << std::endl;
            }
            else
            {
                std::cout << "\tProperties: ";
                auto properties = MakeArray<String>(smartTag->get_Properties()->get_Count());
                int index = 0;

                for (auto cxp : System::IterateOver(smartTag->get_Properties()))
                {
                    properties[index++] = String::Format(u"\"{0}\" = \"{1}\"", cxp->get_Name(), cxp->get_Value());
                }

                std::cout << String::Join(u", ", properties) << std::endl;
            }

            return VisitorAction::Continue;
        }
    };
    //ExEnd

public:
    void TestCreate(SharedPtr<Document> doc)
    {
        auto smartTag = System::DynamicCast<SmartTag>(doc->GetChild(NodeType::SmartTag, 0, true));

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

        smartTag = System::DynamicCast<SmartTag>(doc->GetChild(NodeType::SmartTag, 1, true));

        ASSERT_EQ(u"stockticker", smartTag->get_Element());
        ASSERT_EQ(u"MSFT", smartTag->GetText());
        ASSERT_EQ(u"urn:schemas-microsoft-com:office:smarttags", smartTag->get_Uri());
        ASSERT_EQ(0, smartTag->get_Properties()->get_Count());
    }

    void Properties()
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
        SharedPtr<NodeCollection> smartTags = doc->GetChildNodes(NodeType::SmartTag, true);
        ASSERT_EQ(8, smartTags->get_Count());

        // The last smart tag is of the "Date" type, which we will retrieve here
        auto smartTag = System::DynamicCast<SmartTag>(smartTags->idx_get(7));

        // The Properties attribute, for some smart tags, elaborates on the text object that Word picked up as a smart tag
        // In the case of our "Date" smart tag, its properties will let us know the year, month and day within the smart tag
        SharedPtr<CustomXmlPropertyCollection> properties = smartTag->get_Properties();

        // We can enumerate over the collection and print the aforementioned properties to the console
        ASSERT_EQ(4, properties->get_Count());

        {
            SharedPtr<System::Collections::Generic::IEnumerator<SharedPtr<CustomXmlProperty>>> enumerator = properties->GetEnumerator();
            while (enumerator->MoveNext())
            {
                std::cout << "Property name: " << enumerator->get_Current()->get_Name() << ", value: " << enumerator->get_Current()->get_Value() << std::endl;
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
};

} // namespace ApiExamples
