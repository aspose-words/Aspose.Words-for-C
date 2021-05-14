#pragma once

#include <cstdint>
#include <iostream>
#include <Aspose.Words.Cpp/Body.h>
#include <Aspose.Words.Cpp/Document.h>
#include <Aspose.Words.Cpp/DocumentVisitor.h>
#include <Aspose.Words.Cpp/Markup/CustomXmlProperty.h>
#include <Aspose.Words.Cpp/Markup/CustomXmlPropertyCollection.h>
#include <Aspose.Words.Cpp/Markup/SmartTag.h>
#include <Aspose.Words.Cpp/Node.h>
#include <Aspose.Words.Cpp/NodeCollection.h>
#include <Aspose.Words.Cpp/NodeType.h>
#include <Aspose.Words.Cpp/Paragraph.h>
#include <Aspose.Words.Cpp/Run.h>
#include <Aspose.Words.Cpp/SaveFormat.h>
#include <Aspose.Words.Cpp/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Section.h>
#include <Aspose.Words.Cpp/VisitorAction.h>
#include <system/array.h>
#include <system/collections/ienumerable.h>
#include <system/collections/ienumerator.h>
#include <system/details/dispose_guard.h>
#include <system/enumerator_adapter.h>
#include <system/exceptions.h>
#include <system/linq/enumerable.h>
#include <system/test_tools/compare.h>
#include <system/test_tools/test_tools.h>
#include <testing/test_predicates.h>

#include "ApiExampleBase.h"

using System::ArrayPtr;
using System::MakeArray;
using System::MakeObject;
using System::SharedPtr;
using System::String;

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

        // A smart tag appears in a document with Microsoft Word recognizes a part of its text as some form of data,
        // such as a name, date, or address, and converts it to a hyperlink that displays a purple dotted underline.
        auto smartTag = MakeObject<SmartTag>(doc);

        // Smart tags are composite nodes that contain their recognized text in its entirety.
        // Add contents to this smart tag manually.
        smartTag->AppendChild(MakeObject<Run>(doc, u"May 29, 2019"));

        // Microsoft Word may recognize the above contents as being a date.
        // Smart tags use the "Element" property to reflect the type of data they contain.
        smartTag->set_Element(u"date");

        // Some smart tag types process their contents further into custom XML properties.
        smartTag->get_Properties()->Add(MakeObject<CustomXmlProperty>(u"Day", String::Empty, u"29"));
        smartTag->get_Properties()->Add(MakeObject<CustomXmlProperty>(u"Month", String::Empty, u"5"));
        smartTag->get_Properties()->Add(MakeObject<CustomXmlProperty>(u"Year", String::Empty, u"2019"));

        // Set the smart tag's URI to the default value.
        smartTag->set_Uri(u"urn:schemas-microsoft-com:office:smarttags");

        doc->get_FirstSection()->get_Body()->get_FirstParagraph()->AppendChild(smartTag);
        doc->get_FirstSection()->get_Body()->get_FirstParagraph()->AppendChild(MakeObject<Run>(doc, u" is a date. "));

        // Create another smart tag for a stock ticker.
        smartTag = MakeObject<SmartTag>(doc);
        smartTag->set_Element(u"stockticker");
        smartTag->set_Uri(u"urn:schemas-microsoft-com:office:smarttags");

        smartTag->AppendChild(MakeObject<Run>(doc, u"MSFT"));

        doc->get_FirstSection()->get_Body()->get_FirstParagraph()->AppendChild(smartTag);
        doc->get_FirstSection()->get_Body()->get_FirstParagraph()->AppendChild(MakeObject<Run>(doc, u" is a stock ticker."));

        // Print all the smart tags in our document using a document visitor.
        doc->Accept(MakeObject<ExSmartTag::SmartTagPrinter>());

        // Older versions of Microsoft Word support smart tags.
        doc->Save(ArtifactsDir + u"SmartTag.Create.doc");

        // Use the "RemoveSmartTags" method to remove all smart tags from a document.
        ASSERT_EQ(2, doc->GetChildNodes(NodeType::SmartTag, true)->get_Count());

        doc->RemoveSmartTags();

        ASSERT_EQ(0, doc->GetChildNodes(NodeType::SmartTag, true)->get_Count());
        TestCreate(MakeObject<Document>(ArtifactsDir + u"SmartTag.Create.doc"));
        //ExSkip
    }

    /// <summary>
    /// Prints visited smart tags and their contents.
    /// </summary>
    class SmartTagPrinter : public DocumentVisitor
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
        auto doc = MakeObject<Document>(MyDir + u"Smart tags.doc");

        // A smart tag appears in a document with Microsoft Word recognizes a part of its text as some form of data,
        // such as a name, date, or address, and converts it to a hyperlink that displays a purple dotted underline.
        // In Word 2003, we can enable smart tags via "Tools" -> "AutoCorrect options..." -> "SmartTags".
        // In our input document, there are three objects that Microsoft Word registered as smart tags.
        // Smart tags may be nested, so this collection contains more.
        ArrayPtr<SharedPtr<SmartTag>> smartTags = doc->GetChildNodes(NodeType::SmartTag, true)->LINQ_OfType<SharedPtr<SmartTag>>()->LINQ_ToArray();

        ASSERT_EQ(8, smartTags->get_Length());

        // The "Properties" member of a smart tag contains its metadata, which will be different for each type of smart tag.
        // The properties of a "date"-type smart tag contain its year, month, and day.
        SharedPtr<CustomXmlPropertyCollection> properties = smartTags[7]->get_Properties();

        ASSERT_EQ(4, properties->get_Count());

        {
            SharedPtr<System::Collections::Generic::IEnumerator<SharedPtr<CustomXmlProperty>>> enumerator = properties->GetEnumerator();
            while (enumerator->MoveNext())
            {
                std::cout << "Property name: " << enumerator->get_Current()->get_Name() << ", value: " << enumerator->get_Current()->get_Value() << std::endl;
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
};

} // namespace ApiExamples
