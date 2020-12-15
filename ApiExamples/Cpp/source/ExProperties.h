#pragma once
// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <cstdint>
#include <iostream>
#include <Aspose.Words.Cpp/Layout/Public/LayoutEntityType.h>
#include <Aspose.Words.Cpp/Layout/Public/LayoutEnumerator.h>
#include <Aspose.Words.Cpp/Model/Bookmarks/BookmarkEnd.h>
#include <Aspose.Words.Cpp/Model/Bookmarks/BookmarkStart.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Document/ProtectionType.h>
#include <Aspose.Words.Cpp/Model/Fields/Field.h>
#include <Aspose.Words.Cpp/Model/Fields/FieldCollection.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/LinksAndReferences/FieldHyperlink.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeCollection.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeType.h>
#include <Aspose.Words.Cpp/Model/Properties/BuiltInDocumentProperties.h>
#include <Aspose.Words.Cpp/Model/Properties/CustomDocumentProperties.h>
#include <Aspose.Words.Cpp/Model/Properties/DocumentProperty.h>
#include <Aspose.Words.Cpp/Model/Properties/DocumentSecurity.h>
#include <Aspose.Words.Cpp/Model/Properties/PropertyType.h>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Model/Settings/WriteProtection.h>
#include <Aspose.Words.Cpp/Model/Text/Range.h>
#include <system/array.h>
#include <system/collections/icollection.h>
#include <system/collections/ienumerator.h>
#include <system/convert.h>
#include <system/date_time.h>
#include <system/details/dispose_guard.h>
#include <system/enum.h>
#include <system/enum_helpers.h>
#include <system/enumerator_adapter.h>
#include <system/exceptions.h>
#include <system/io/file.h>
#include <system/io/file_mode.h>
#include <system/io/file_stream.h>
#include <system/object_ext.h>
#include <system/scope_guard.h>
#include <system/test_tools/compare.h>
#include <system/test_tools/test_tools.h>
#include <system/timespan.h>
#include <testing/test_predicates.h>

#include "ApiExampleBase.h"
#include "DocumentHelper.h"
#include "TestUtil.h"

using System::String;
using System::SharedPtr;
using System::ArrayPtr;
using System::MakeObject;
using System::MakeArray;

using namespace Aspose::Words;
using namespace Aspose::Words::Fields;
using namespace Aspose::Words::Layout;
using namespace Aspose::Words::Properties;

namespace ApiExamples {

class ExProperties : public ApiExampleBase
{
public:
    void BuiltIn()
    {
        //ExStart
        //ExFor:BuiltInDocumentProperties
        //ExFor:Document.BuiltInDocumentProperties
        //ExFor:Document.CustomDocumentProperties
        //ExFor:DocumentProperty
        //ExFor:DocumentProperty.Name
        //ExFor:DocumentProperty.Value
        //ExFor:DocumentProperty.Type
        //ExSummary:Shows how to work with built in document properties.
        auto doc = MakeObject<Document>(MyDir + u"Properties.docx");

        // Some information about the document is stored in member attributes, and can be accessed like this
        std::cout << "Document filename:\n\t \"" << doc->get_OriginalFileName() << "\"" << std::endl;

        // Most of the document's metadata, such as author name, file size,
        // word/page counts can be found in the built-in properties collection like this
        std::cout << "Built-in Properties:" << std::endl;
        for (auto docProperty : System::IterateOver(doc->get_BuiltInDocumentProperties()))
        {
            std::cout << docProperty->get_Name() << std::endl;
            std::cout << String::Format(u"\tType:\t{0}", docProperty->get_Type()) << std::endl;

            // Some properties may store multiple values
            if (System::ObjectExt::Is<System::Collections::Generic::ICollection<SharedPtr<System::Object>>>(docProperty->get_Value()))
            {
                for (auto value : System::IterateOver(
                         System::DynamicCast_noexcept<System::Collections::Generic::ICollection<SharedPtr<System::Object>>>(docProperty->get_Value())))
                {
                    std::cout << "\tValue:\t\"" << value << "\"" << std::endl;
                }
            }
            else
            {
                std::cout << "\tValue:\t\"" << docProperty->get_Value() << "\"" << std::endl;
            }
        }
        //ExEnd

        ASSERT_EQ(28, doc->get_BuiltInDocumentProperties()->get_Count());
    }

    void Custom()
    {
        //ExStart
        //ExFor:BuiltInDocumentProperties.Item(String)
        //ExFor:CustomDocumentProperties
        //ExFor:DocumentProperty.ToString
        //ExFor:DocumentPropertyCollection.Count
        //ExFor:DocumentPropertyCollection.Item(int)
        //ExSummary:Shows how to work with custom document properties.
        auto doc = MakeObject<Document>(MyDir + u"Properties.docx");

        // A document's built-in properties contain a set of predetermined keys
        // with values such as the author's name or document's word count
        // We can add our own keys and values to a custom properties collection also
        // Before we add a custom property, we need to make sure that one with the same name does not already exist
        ASSERT_EQ(u"Value of custom document property", System::ObjectExt::ToString(doc->get_CustomDocumentProperties()->idx_get(u"CustomProperty")));

        doc->get_CustomDocumentProperties()->Add(u"CustomProperty2", String(u"Value of custom document property #2"));

        // Iterate over all the custom document properties
        std::cout << "Custom Properties:" << std::endl;
        for (auto customDocumentProperty : System::IterateOver(doc->get_CustomDocumentProperties()))
        {
            std::cout << customDocumentProperty->get_Name() << std::endl;
            std::cout << String::Format(u"\tType:\t{0}", customDocumentProperty->get_Type()) << std::endl;
            std::cout << "\tValue:\t\"" << customDocumentProperty->get_Value() << "\"" << std::endl;
        }
        //ExEnd

        ASSERT_EQ(2, doc->get_CustomDocumentProperties()->get_Count());
    }

    void Description()
    {
        //ExStart
        //ExFor:BuiltInDocumentProperties.Author
        //ExFor:BuiltInDocumentProperties.Category
        //ExFor:BuiltInDocumentProperties.Comments
        //ExFor:BuiltInDocumentProperties.Keywords
        //ExFor:BuiltInDocumentProperties.Subject
        //ExFor:BuiltInDocumentProperties.Title
        //ExSummary:Shows how to work with document properties in the "Description" category.
        auto doc = MakeObject<Document>();

        // The properties we will work with are members of the BuiltInDocumentProperties attribute
        SharedPtr<BuiltInDocumentProperties> properties = doc->get_BuiltInDocumentProperties();

        // Set the values of some descriptive properties
        // These are metadata that can be glanced at without opening the document in the "Details" or "Content" folder views in Windows Explorer
        // The "Details" view has columns dedicated to these properties
        // Fields such as AUTHOR, SUBJECT, TITLE etc. can be used to display these values inside the document
        properties->set_Author(u"John Doe");
        properties->set_Title(u"John's Document");
        properties->set_Subject(u"My subject");
        properties->set_Category(u"My category");
        properties->set_Comments(String::Format(u"This is {0}'s document about {1}", properties->get_Author(), properties->get_Subject()));

        // Tags can be used as keywords and are separated by semicolons
        properties->set_Keywords(u"Tag 1; Tag 2; Tag 3");

        // When right clicking the document file in Windows Explorer, these properties are found in Properties > Details > Description
        doc->Save(ArtifactsDir + u"Properties.Description.docx");
        //ExEnd

        properties = MakeObject<Document>(ArtifactsDir + u"Properties.Description.docx")->get_BuiltInDocumentProperties();

        ASSERT_EQ(u"John Doe", properties->get_Author());
        ASSERT_EQ(u"My category", properties->get_Category());
        ASSERT_EQ(String::Format(u"This is {0}'s document about {1}", properties->get_Author(), properties->get_Subject()), properties->get_Comments());
        ASSERT_EQ(u"Tag 1; Tag 2; Tag 3", properties->get_Keywords());
        ASSERT_EQ(u"My subject", properties->get_Subject());
        ASSERT_EQ(u"John's Document", properties->get_Title());
    }

    void Origin()
    {
        //ExStart
        //ExFor:BuiltInDocumentProperties.Company
        //ExFor:BuiltInDocumentProperties.CreatedTime
        //ExFor:BuiltInDocumentProperties.LastPrinted
        //ExFor:BuiltInDocumentProperties.LastSavedBy
        //ExFor:BuiltInDocumentProperties.LastSavedTime
        //ExFor:BuiltInDocumentProperties.Manager
        //ExFor:BuiltInDocumentProperties.NameOfApplication
        //ExFor:BuiltInDocumentProperties.RevisionNumber
        //ExFor:BuiltInDocumentProperties.Template
        //ExFor:BuiltInDocumentProperties.TotalEditingTime
        //ExFor:BuiltInDocumentProperties.Version
        //ExSummary:Shows how to work with document properties in the "Origin" category.
        auto doc = MakeObject<Document>(MyDir + u"Properties.docx");

        // The properties we will work with are members of the BuiltInDocumentProperties attribute
        SharedPtr<BuiltInDocumentProperties> properties = doc->get_BuiltInDocumentProperties();

        // Since this document has been edited and printed in the past, values generated by Microsoft Word will appear here
        // These values can be glanced at by right clicking the file in Windows Explorer, without actually opening the document
        // Fields such as PRINTDATE, EDITTIME etc. can display these values inside the document
        std::cout << "Created using " << properties->get_NameOfApplication() << ", on " << properties->get_CreatedTime() << std::endl;
        std::cout << "Minutes spent editing: " << properties->get_TotalEditingTime() << std::endl;
        std::cout << "Date/time last printed: " << properties->get_LastPrinted() << std::endl;
        std::cout << "Template document: " << properties->get_Template() << std::endl;

        // We can set these properties ourselves
        properties->set_Company(u"Doe Ltd.");
        properties->set_Manager(u"Jane Doe");
        properties->set_Version(5);
        properties->set_RevisionNumber(properties->get_RevisionNumber() + 1);

        // If we plan on programmatically saving the document, we may record some details like this
        properties->set_LastSavedBy(u"John Doe");
        properties->set_LastSavedTime(System::DateTime::get_Now());

        // When right clicking the document file in Windows Explorer, these properties are found in Properties > Details > Origin
        doc->Save(ArtifactsDir + u"Properties.Origin.docx");
        //ExEnd

        properties = MakeObject<Document>(ArtifactsDir + u"Properties.Origin.docx")->get_BuiltInDocumentProperties();

        ASSERT_EQ(u"Doe Ltd.", properties->get_Company());
        ASSERT_EQ(System::DateTime(2006, 4, 25, 10, 10, 0), properties->get_CreatedTime());
        ASSERT_EQ(System::DateTime(2019, 4, 21, 10, 0, 0), properties->get_LastPrinted());
        ASSERT_EQ(u"John Doe", properties->get_LastSavedBy());
        TestUtil::VerifyDate(System::DateTime::get_Now(), properties->get_LastSavedTime(), System::TimeSpan::FromSeconds(5));
        ASSERT_EQ(u"Jane Doe", properties->get_Manager());
        ASSERT_EQ(u"Microsoft Office Word", properties->get_NameOfApplication());
        ASSERT_EQ(12, properties->get_RevisionNumber());
        ASSERT_EQ(u"Normal", properties->get_Template());
        ASSERT_EQ(8, properties->get_TotalEditingTime());
        ASSERT_EQ(786432, properties->get_Version());
    }

    //ExStart
    //ExFor:BuiltInDocumentProperties.Bytes
    //ExFor:BuiltInDocumentProperties.Characters
    //ExFor:BuiltInDocumentProperties.CharactersWithSpaces
    //ExFor:BuiltInDocumentProperties.ContentStatus
    //ExFor:BuiltInDocumentProperties.ContentType
    //ExFor:BuiltInDocumentProperties.Lines
    //ExFor:BuiltInDocumentProperties.LinksUpToDate
    //ExFor:BuiltInDocumentProperties.Pages
    //ExFor:BuiltInDocumentProperties.Paragraphs
    //ExFor:BuiltInDocumentProperties.Words
    //ExSummary:Shows how to work with document properties in the "Content" category.
    void Content()
    {
        // Open a document with a couple paragraphs of content
        auto doc = MakeObject<Document>(MyDir + u"Paragraphs.docx");

        // The properties we will work with are members of the BuiltInDocumentProperties attribute
        SharedPtr<BuiltInDocumentProperties> properties = doc->get_BuiltInDocumentProperties();

        // By using built in properties,
        // we can treat document statistics such as word/page/character counts as metadata that can be glanced at without opening the document
        // These properties are accessed by right-clicking the file in Windows Explorer and navigating to Properties > Details > Content
        // If we want to display this data inside the document, we can use fields such as NUMPAGES, NUMWORDS, NUMCHARS etc.
        // Also, these values can also be viewed in Microsoft Word by navigating File > Properties > Advanced Properties > Statistics
        // Page count: The PageCount attribute shows the page count in real time and its value can be assigned to the Pages property
        properties->set_Pages(doc->get_PageCount());
        ASSERT_EQ(6, properties->get_Pages());

        // Word count: The UpdateWordCount() automatically assigns the real time word/character counts to the respective built in properties
        doc->UpdateWordCount();
        ASSERT_EQ(1035, properties->get_Words());
        ASSERT_EQ(6026, properties->get_Characters());
        ASSERT_EQ(7041, properties->get_CharactersWithSpaces());

        // Line count: Count the lines in a document and assign value to the Lines property
        auto lineCounter = MakeObject<ExProperties::LineCounter>(doc);
        properties->set_Lines(lineCounter->GetLineCount());
        ASSERT_EQ(142, properties->get_Lines());

        // Paragraph count: Assign the size of the count of child Paragraph-nodes to the Paragraphs built in property
        properties->set_Paragraphs(doc->GetChildNodes(NodeType::Paragraph, true)->get_Count());
        ASSERT_EQ(29, properties->get_Paragraphs());

        // Check the real file size of our document
        ASSERT_EQ(20310, properties->get_Bytes());

        // Template: The Template attribute can reflect the filename of the attached template document
        doc->set_AttachedTemplate(MyDir + u"Business brochure.dotx");
        ASSERT_EQ(u"Normal", properties->get_Template());
        properties->set_Template(doc->get_AttachedTemplate());

        // Content status: This is a descriptive field
        properties->set_ContentStatus(u"Draft");

        // Content type: Upon saving, any value we assign to this field will be overwritten by the MIME type of the output save format
        ASSERT_EQ(String::Empty, properties->get_ContentType());

        // If the document contains links and they are all up to date, we can set this to true
        ASSERT_FALSE(properties->get_LinksUpToDate());

        doc->Save(ArtifactsDir + u"Properties.Content.docx");
        TestContent(MakeObject<Document>(ArtifactsDir + u"Properties.Content.docx"));
        //ExSkip
    }

private:
    /// <summary>
    /// Util class that counts the lines in a document.
    /// Upon construction, traverses the document's layout entities tree,
    /// counting entities of the "Line" type that also contain real text.
    /// </summary>
    class LineCounter : public System::Object
    {
    public:
        LineCounter(SharedPtr<Document> doc) : mLineCount(0), mScanningLineForRealText(false)
        {
            // Self Reference+++++++++++++++++++++++++++++++++++++++++++++++++++++++++
            System::Details::ThisProtector __local_self_ref(this);
            //---------------------------------------------------------Self Reference

            mLayoutEnumerator = MakeObject<LayoutEnumerator>(doc);

            CountLines();
        }

        int GetLineCount()
        {
            return mLineCount;
        }

    private:
        SharedPtr<LayoutEnumerator> mLayoutEnumerator;
        int mLineCount;
        bool mScanningLineForRealText;

        void CountLines()
        {
            do
            {
                if (mLayoutEnumerator->get_Type() == LayoutEntityType::Line)
                {
                    mScanningLineForRealText = true;
                }

                if (mLayoutEnumerator->MoveFirstChild())
                {
                    if (mScanningLineForRealText && mLayoutEnumerator->get_Kind().StartsWith(u"TEXT"))
                    {
                        mLineCount++;
                        mScanningLineForRealText = false;
                    }
                    CountLines();
                    mLayoutEnumerator->MoveParent();
                }
            } while (mLayoutEnumerator->MoveNext());
        }
    };
    //ExEnd

protected:
    void TestContent(SharedPtr<Document> doc)
    {
        SharedPtr<BuiltInDocumentProperties> properties = doc->get_BuiltInDocumentProperties();

        ASSERT_EQ(6, properties->get_Pages());

        ASSERT_EQ(1035, properties->get_Words());
        ASSERT_EQ(6026, properties->get_Characters());
        ASSERT_EQ(7041, properties->get_CharactersWithSpaces());
        ASSERT_EQ(142, properties->get_Lines());
        ASSERT_EQ(29, properties->get_Paragraphs());
        ASSERT_NEAR(15500, properties->get_Bytes(), 200);
        ASSERT_EQ(MyDir.Replace(u"\\\\", u"\\") + u"Business brochure.dotx", properties->get_Template());
        ASSERT_EQ(u"Draft", properties->get_ContentStatus());
        ASSERT_EQ(String::Empty, properties->get_ContentType());
        ASSERT_FALSE(properties->get_LinksUpToDate());
    }

public:
    void Thumbnail()
    {
        //ExStart
        //ExFor:BuiltInDocumentProperties.Thumbnail
        //ExFor:DocumentProperty.ToByteArray
        //ExSummary:Shows how to append a thumbnail to an Epub document.
        // Create a blank document and add some text with a DocumentBuilder
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);
        builder->Writeln(u"Hello world!");

        // The thumbnail property resides in a document's built-in properties, but is used exclusively by Epub e-book documents
        SharedPtr<BuiltInDocumentProperties> properties = doc->get_BuiltInDocumentProperties();

        // Load an image from our file system into a byte array
        ArrayPtr<uint8_t> thumbnailBytes = System::IO::File::ReadAllBytes(ImageDir + u"Logo.jpg");

        // Set the value of the Thumbnail property to the array from above
        properties->set_Thumbnail(thumbnailBytes);

        // Our thumbnail should be visible at the start of the document, before the text we added
        doc->Save(ArtifactsDir + u"Properties.Thumbnail.epub");

        // We can also extract a thumbnail property into a byte array and then into the local file system like this
        SharedPtr<DocumentProperty> thumbnail = doc->get_BuiltInDocumentProperties()->idx_get(u"Thumbnail");
        System::IO::File::WriteAllBytes(ArtifactsDir + u"Properties.Thumbnail.gif", thumbnail->ToByteArray());
        //ExEnd

        {
            auto imgStream = MakeObject<System::IO::FileStream>(ArtifactsDir + u"Properties.Thumbnail.gif", System::IO::FileMode::Open);
            TestUtil::VerifyImage(400, 400, imgStream);
        }
    }

    void HyperlinkBase()
    {
        //ExStart
        //ExFor:BuiltInDocumentProperties.HyperlinkBase
        //ExSummary:Shows how to store the base part of a hyperlink in the document's properties.
        // Create a blank document and a DocumentBuilder
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Insert a relative hyperlink to "Document.docx", which will open that document when clicked on
        builder->InsertHyperlink(u"Relative hyperlink", u"Document.docx", false);

        // If we do not have a "Document.docx" in the same folder as the document we are about to save, we will end up with a broken link
        ASSERT_FALSE(System::IO::File::Exists(ArtifactsDir + u"Document.docx"));
        doc->Save(ArtifactsDir + u"Properties.HyperlinkBase.BrokenLink.docx");

        // We could keep prepending something like "C:\users\...\data" to every hyperlink we place to remedy this
        // Alternatively, if we know that all our linked files will come from the same folder,
        // we could set a base hyperlink in the document properties, keeping our hyperlinks short
        SharedPtr<BuiltInDocumentProperties> properties = doc->get_BuiltInDocumentProperties();
        properties->set_HyperlinkBase(MyDir);

        ASSERT_TRUE(System::IO::File::Exists(properties->get_HyperlinkBase() +
                                             (System::DynamicCast<FieldHyperlink>(doc->get_Range()->get_Fields()->idx_get(0)))->get_Address()));

        doc->Save(ArtifactsDir + u"Properties.HyperlinkBase.WorkingLink.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Properties.HyperlinkBase.BrokenLink.docx");
        properties = doc->get_BuiltInDocumentProperties();

        ASSERT_EQ(String::Empty, properties->get_HyperlinkBase());

        doc = MakeObject<Document>(ArtifactsDir + u"Properties.HyperlinkBase.WorkingLink.docx");
        properties = doc->get_BuiltInDocumentProperties();

        ASSERT_EQ(MyDir, properties->get_HyperlinkBase());
        ASSERT_TRUE(System::IO::File::Exists(properties->get_HyperlinkBase() +
                                             (System::DynamicCast<FieldHyperlink>(doc->get_Range()->get_Fields()->idx_get(0)))->get_Address()));
    }

    void HeadingPairs()
    {
        //ExStart
        //ExFor:Properties.BuiltInDocumentProperties.HeadingPairs
        //ExFor:Properties.BuiltInDocumentProperties.TitlesOfParts
        //ExSummary:Shows the relationship between HeadingPairs and TitlesOfParts properties.
        // Open a document that contains entries in the HeadingPairs/TitlesOfParts properties
        auto doc = MakeObject<Document>(MyDir + u"Heading pairs and titles of parts.docx");

        // We can find the combined values of these collections in File > Properties > Advanced Properties > Contents tab
        // The HeadingPairs property is a collection of <string, int> pairs that determines
        // how many document parts a heading spans over
        ArrayPtr<SharedPtr<System::Object>> headingPairs = doc->get_BuiltInDocumentProperties()->get_HeadingPairs();

        // The TitlesOfParts property contains the names of parts that belong to the above headings
        ArrayPtr<String> titlesOfParts = doc->get_BuiltInDocumentProperties()->get_TitlesOfParts();

        int headingPairsIndex = 0;
        int titlesOfPartsIndex = 0;
        while (headingPairsIndex < headingPairs->get_Length())
        {
            std::cout << "Parts for " << headingPairs[headingPairsIndex++] << ":" << std::endl;
            int partsCount = System::Convert::ToInt32(headingPairs[headingPairsIndex++]);

            for (int i = 0; i < partsCount; i++)
            {
                std::cout << "\t\"" << titlesOfParts[titlesOfPartsIndex++] << "\"" << std::endl;
            }
        }
        //ExEnd

        // There are 6 array elements designating 3 heading/part count pairs
        ASSERT_EQ(6, headingPairs->get_Length());
        ASSERT_EQ(u"Title", System::ObjectExt::ToString(headingPairs[0]));
        ASSERT_EQ(u"1", System::ObjectExt::ToString(headingPairs[1]));
        ASSERT_EQ(u"Heading 1", System::ObjectExt::ToString(headingPairs[2]));
        ASSERT_EQ(u"5", System::ObjectExt::ToString(headingPairs[3]));
        ASSERT_EQ(u"Heading 2", System::ObjectExt::ToString(headingPairs[4]));
        ASSERT_EQ(u"2", System::ObjectExt::ToString(headingPairs[5]));

        ASSERT_EQ(8, titlesOfParts->get_Length());
        // "Title"
        ASSERT_EQ(u"", titlesOfParts[0]);
        // "Heading 1"
        ASSERT_EQ(u"Part1", titlesOfParts[1]);
        ASSERT_EQ(u"Part2", titlesOfParts[2]);
        ASSERT_EQ(u"Part3", titlesOfParts[3]);
        ASSERT_EQ(u"Part4", titlesOfParts[4]);
        ASSERT_EQ(u"Part5", titlesOfParts[5]);
        // "Heading 2"
        ASSERT_EQ(u"Part6", titlesOfParts[6]);
        ASSERT_EQ(u"Part7", titlesOfParts[7]);
    }

    void Security()
    {
        //ExStart
        //ExFor:Properties.BuiltInDocumentProperties.Security
        //ExFor:Properties.DocumentSecurity
        //ExSummary:Shows how to use document properties to display the security level of a document.
        auto doc = MakeObject<Document>();

        // The "Security" property serves as a description of the security level of a document
        ASSERT_EQ(DocumentSecurity::None, doc->get_BuiltInDocumentProperties()->get_Security());

        // Upon saving a document after setting its security level, Aspose automatically updates this property to the appropriate value
        doc->get_WriteProtection()->set_ReadOnlyRecommended(true);
        doc->Save(ArtifactsDir + u"Properties.Security.ReadOnlyRecommended.docx");

        // Open a document and verify its security level
        ASSERT_EQ(
            DocumentSecurity::ReadOnlyRecommended,
            MakeObject<Document>(ArtifactsDir + u"Properties.Security.ReadOnlyRecommended.docx")->get_BuiltInDocumentProperties()->get_Security());

        // Create a new document and set it to Write-Protected
        doc = MakeObject<Document>();

        ASSERT_FALSE(doc->get_WriteProtection()->get_IsWriteProtected());
        doc->get_WriteProtection()->SetPassword(u"MyPassword");
        ASSERT_TRUE(doc->get_WriteProtection()->ValidatePassword(u"MyPassword"));
        ASSERT_TRUE(doc->get_WriteProtection()->get_IsWriteProtected());
        doc->Save(ArtifactsDir + u"Properties.Security.ReadOnlyEnforced.docx");

        // This document's security level counts as "ReadOnlyEnforced"
        ASSERT_EQ(DocumentSecurity::ReadOnlyEnforced,
                  MakeObject<Document>(ArtifactsDir + u"Properties.Security.ReadOnlyEnforced.docx")->get_BuiltInDocumentProperties()->get_Security());

        // Since this is still a descriptive property, we can protect a document and pick a suitable value ourselves
        doc = MakeObject<Document>();

        doc->Protect(ProtectionType::AllowOnlyComments, u"MyPassword");
        doc->get_BuiltInDocumentProperties()->set_Security(DocumentSecurity::ReadOnlyExceptAnnotations);
        doc->Save(ArtifactsDir + u"Properties.Security.ReadOnlyExceptAnnotations.docx");

        ASSERT_EQ(DocumentSecurity::ReadOnlyExceptAnnotations,
                  MakeObject<Document>(ArtifactsDir + u"Properties.Security.ReadOnlyExceptAnnotations.docx")
                      ->get_BuiltInDocumentProperties()
                      ->get_Security());
        //ExEnd
    }

    void CustomNamedAccess()
    {
        //ExStart
        //ExFor:DocumentPropertyCollection.Item(String)
        //ExFor:CustomDocumentProperties.Add(String,DateTime)
        //ExFor:DocumentProperty.ToDateTime
        //ExSummary:Shows how to create a custom document property with the value of a date and time.
        auto doc = MakeObject<Document>();

        doc->get_CustomDocumentProperties()->Add(u"AuthorizedDate", System::DateTime::get_Now());

        std::cout << "Document authorized on " << doc->get_CustomDocumentProperties()->idx_get(u"AuthorizedDate")->ToDateTime() << std::endl;
        //ExEnd

        TestUtil::VerifyDate(System::DateTime::get_Now(),
                             DocumentHelper::SaveOpen(doc)->get_CustomDocumentProperties()->idx_get(u"AuthorizedDate")->ToDateTime(),
                             System::TimeSpan::FromSeconds(1));
    }

    void LinkCustomDocumentPropertiesToBookmark()
    {
        //ExStart
        //ExFor:CustomDocumentProperties.AddLinkToContent(String, String)
        //ExFor:DocumentProperty.IsLinkToContent
        //ExFor:DocumentProperty.LinkSource
        //ExSummary:Shows how to link a custom document property to a bookmark.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);
        builder->StartBookmark(u"MyBookmark");
        builder->Write(u"MyBookmark contents.");
        builder->EndBookmark(u"MyBookmark");

        // Add linked to content property
        SharedPtr<CustomDocumentProperties> customProperties = doc->get_CustomDocumentProperties();
        SharedPtr<DocumentProperty> customProperty = customProperties->AddLinkToContent(u"Bookmark", u"MyBookmark");

        // Check whether the property is linked to content
        ASPOSE_ASSERT_EQ(true, customProperty->get_IsLinkToContent());
        ASSERT_EQ(u"MyBookmark", customProperty->get_LinkSource());
        ASPOSE_ASSERT_EQ(u"MyBookmark contents.", customProperty->get_Value());

        doc->Save(ArtifactsDir + u"Properties.LinkCustomDocumentPropertiesToBookmark.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Properties.LinkCustomDocumentPropertiesToBookmark.docx");
        customProperty = doc->get_CustomDocumentProperties()->idx_get(u"Bookmark");

        ASPOSE_ASSERT_EQ(true, customProperty->get_IsLinkToContent());
        ASSERT_EQ(u"MyBookmark", customProperty->get_LinkSource());
        ASPOSE_ASSERT_EQ(u"MyBookmark contents.", customProperty->get_Value());
    }

    void DocumentPropertyCollection()
    {
        //ExStart
        //ExFor:CustomDocumentProperties.Add(String,String)
        //ExFor:CustomDocumentProperties.Add(String,Boolean)
        //ExFor:CustomDocumentProperties.Add(String,int)
        //ExFor:CustomDocumentProperties.Add(String,DateTime)
        //ExFor:CustomDocumentProperties.Add(String,Double)
        //ExFor:DocumentProperty.Type
        //ExFor:Properties.DocumentPropertyCollection
        //ExFor:Properties.DocumentPropertyCollection.Clear
        //ExFor:Properties.DocumentPropertyCollection.Contains(System.String)
        //ExFor:Properties.DocumentPropertyCollection.GetEnumerator
        //ExFor:Properties.DocumentPropertyCollection.IndexOf(System.String)
        //ExFor:Properties.DocumentPropertyCollection.RemoveAt(System.Int32)
        //ExFor:Properties.DocumentPropertyCollection.Remove
        //ExFor:PropertyType
        //ExSummary:Shows how to add custom properties to a document.
        auto doc = MakeObject<Document>();
        SharedPtr<CustomDocumentProperties> properties = doc->get_CustomDocumentProperties();

        // The custom property collection will be empty by default
        ASSERT_EQ(0, properties->get_Count());

        // We can populate it with key/value pairs with a variety of value types
        properties->Add(u"Authorized", true);
        properties->Add(u"Authorized By", String(u"John Doe"));
        properties->Add(u"Authorized Date", System::DateTime::get_Today());
        properties->Add(u"Authorized Revision", doc->get_BuiltInDocumentProperties()->get_RevisionNumber());
        properties->Add(u"Authorized Amount", 123.45);

        // Custom properties are automatically sorted in alphabetic order
        ASSERT_EQ(1, properties->IndexOf(u"Authorized Amount"));
        ASSERT_EQ(5, properties->get_Count());

        // Enumerate and print all custom properties
        {
            SharedPtr<System::Collections::Generic::IEnumerator<SharedPtr<DocumentProperty>>> enumerator = properties->GetEnumerator();
            while (enumerator->MoveNext())
            {
                std::cout << String::Format(u"Name: \"{0}\"\n\tType: \"{1}\"\n\tValue: \"{2}\"", enumerator->get_Current()->get_Name(),
                                                    enumerator->get_Current()->get_Type(), enumerator->get_Current()->get_Value())
                          << std::endl;
            }
        }

        // We can view/edit custom properties by opening the document and looking in File > Properties > Advanced Properties > Custom
        doc->Save(ArtifactsDir + u"Properties.DocumentPropertyCollection.docx");

        // We can remove elements from the property collection by index or by name
        properties->RemoveAt(1);
        ASSERT_FALSE(properties->Contains(u"Authorized Amount"));
        ASSERT_EQ(4, properties->get_Count());

        properties->Remove(u"Authorized Revision");
        ASSERT_FALSE(properties->Contains(u"Authorized Revision"));
        ASSERT_EQ(3, properties->get_Count());

        // We can also empty the entire custom property collection at once
        properties->Clear();
        ASSERT_EQ(0, properties->get_Count());
        //ExEnd
    }

    void PropertyTypes()
    {
        //ExStart
        //ExFor:DocumentProperty.ToBool
        //ExFor:DocumentProperty.ToInt
        //ExFor:DocumentProperty.ToDouble
        //ExFor:DocumentProperty.ToString
        //ExFor:DocumentProperty.ToDateTime
        //ExSummary:Shows various type conversion methods of custom document properties.
        auto doc = MakeObject<Document>();
        SharedPtr<CustomDocumentProperties> properties = doc->get_CustomDocumentProperties();

        System::DateTime authDate = System::DateTime::get_Today();
        properties->Add(u"Authorized", true);
        properties->Add(u"Authorized By", String(u"John Doe"));
        properties->Add(u"Authorized Date", authDate);
        properties->Add(u"Authorized Revision", doc->get_BuiltInDocumentProperties()->get_RevisionNumber());
        properties->Add(u"Authorized Amount", 123.45);

        ASPOSE_ASSERT_EQ(true, properties->idx_get(u"Authorized")->ToBool());
        ASSERT_EQ(u"John Doe", System::ObjectExt::ToString(properties->idx_get(u"Authorized By")));
        ASSERT_EQ(authDate, properties->idx_get(u"Authorized Date")->ToDateTime());
        ASSERT_EQ(1, properties->idx_get(u"Authorized Revision")->ToInt());
        ASPOSE_ASSERT_EQ(123.45, properties->idx_get(u"Authorized Amount")->ToDouble());
        //ExEnd
    }
};

} // namespace ApiExamples
