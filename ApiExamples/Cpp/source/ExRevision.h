#pragma once
// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <cstdint>
#include <iostream>
#include <Aspose.Words.Cpp/Layout/Public/LayoutOptions.h>
#include <Aspose.Words.Cpp/Layout/Public/RevisionColor.h>
#include <Aspose.Words.Cpp/Layout/Public/RevisionOptions.h>
#include <Aspose.Words.Cpp/Layout/Public/RevisionTextEffect.h>
#include <Aspose.Words.Cpp/Layout/Public/ShowInBalloons.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Nodes/Node.h>
#include <Aspose.Words.Cpp/Model/Revisions/Revision.h>
#include <Aspose.Words.Cpp/Model/Revisions/RevisionCollection.h>
#include <Aspose.Words.Cpp/Model/Revisions/RevisionGroup.h>
#include <Aspose.Words.Cpp/Model/Revisions/RevisionGroupCollection.h>
#include <Aspose.Words.Cpp/Model/Revisions/RevisionType.h>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Model/Sections/Body.h>
#include <Aspose.Words.Cpp/Model/Sections/Section.h>
#include <Aspose.Words.Cpp/Model/Styles/Style.h>
#include <Aspose.Words.Cpp/Model/Text/Paragraph.h>
#include <Aspose.Words.Cpp/Model/Text/Run.h>
#include <Aspose.Words.Cpp/Model/Text/RunCollection.h>
#include <system/collections/ienumerator.h>
#include <system/date_time.h>
#include <system/details/dispose_guard.h>
#include <system/enum.h>
#include <system/enumerator_adapter.h>
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
using namespace Aspose::Words::Layout;

namespace ApiExamples {

class ExRevision : public ApiExampleBase
{
public:
    void Revisions()
    {
        //ExStart
        //ExFor:Revision
        //ExFor:Revision.Accept
        //ExFor:Revision.Author
        //ExFor:Revision.DateTime
        //ExFor:Revision.Group
        //ExFor:Revision.Reject
        //ExFor:Revision.RevisionType
        //ExFor:RevisionCollection
        //ExFor:RevisionCollection.Item(Int32)
        //ExFor:RevisionCollection.Count
        //ExFor:RevisionType
        //ExFor:Document.HasRevisions
        //ExFor:Document.TrackRevisions
        //ExFor:Document.Revisions
        //ExSummary:Shows how to work with revisions in a document.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Standard editing of the document does not count as a revision.
        builder->Write(u"This does not count as a revision. ");

        ASSERT_FALSE(doc->get_HasRevisions());

        // To register our edits as revisions, we need to declare an author, and then start tracking them.
        doc->StartTrackRevisions(u"John Doe", System::DateTime::get_Now());

        builder->Write(u"This is revision #1. ");

        ASSERT_TRUE(doc->get_HasRevisions());
        ASSERT_EQ(1, doc->get_Revisions()->get_Count());

        // This flag corresponds to the Review -> Tracking -> "Track Changes" option is turned on in Microsoft Word,
        // and it is independent of the programmatic revision tracking that is taking place here.
        ASSERT_FALSE(doc->get_TrackRevisions());

        // Our first revision is an insertion-type revision since we added text with the document builder.
        SharedPtr<Revision> revision = doc->get_Revisions()->idx_get(0);
        ASSERT_EQ(u"John Doe", revision->get_Author());
        ASSERT_EQ(u"This is revision #1. ", revision->get_ParentNode()->GetText());
        ASSERT_EQ(RevisionType::Insertion, revision->get_RevisionType());
        ASSERT_EQ(revision->get_DateTime().get_Date(), System::DateTime::get_Now().get_Date());
        ASPOSE_ASSERT_EQ(doc->get_Revisions()->get_Groups()->idx_get(0), revision->get_Group());

        // Remove a run to create a deletion-type revision.
        doc->get_FirstSection()->get_Body()->get_FirstParagraph()->get_Runs()->idx_get(0)->Remove();

        // Adding a new revision places it at the beginning of the revision collection.
        ASSERT_EQ(RevisionType::Deletion, doc->get_Revisions()->idx_get(0)->get_RevisionType());
        ASSERT_EQ(2, doc->get_Revisions()->get_Count());

        // Insert revisions are treated as document text by the GetText() method before they are accepted
        // since they are still nodes with text and are in the body.
        ASSERT_EQ(u"This does not count as a revision. This is revision #1.", doc->GetText().Trim());

        // Accepting the delete revision will remove its parent node from the paragraph text,
        // and then remove the revision itself from the collection.
        doc->get_Revisions()->idx_get(0)->Accept();

        ASSERT_EQ(1, doc->get_Revisions()->get_Count());

        // Accepting a delete revision removes all the nodes that it concerns,
        // so their contents will no longer be anywhere in the document.
        ASSERT_EQ(u"This is revision #1.", doc->GetText().Trim());

        // The insertion-type revision is now at index 0, which we can reject to ignore and discard it.
        doc->get_Revisions()->idx_get(0)->Reject();

        ASSERT_EQ(0, doc->get_Revisions()->get_Count());
        ASSERT_EQ(u"", doc->GetText().Trim());
        //ExEnd
    }

    void RevisionCollection_()
    {
        //ExStart
        //ExFor:Revision.ParentStyle
        //ExFor:RevisionCollection.GetEnumerator
        //ExFor:RevisionCollection.Groups
        //ExFor:RevisionCollection.RejectAll
        //ExFor:RevisionGroupCollection.GetEnumerator
        //ExSummary:Shows how to iterate through a document's revisions.
        auto doc = MakeObject<Document>(MyDir + u"Revisions.docx");
        SharedPtr<RevisionCollection> revisions = doc->get_Revisions();

        // This collection itself has a collection of revision groups.
        // Each group is a sequence of adjacent revisions.
        ASSERT_EQ(7, revisions->get_Groups()->get_Count());
        //ExSkip
        std::cout << revisions->get_Groups()->get_Count() << " revision groups:" << std::endl;

        // Iterate over the collection of groups and print the text that the revision concerns.
        {
            SharedPtr<System::Collections::Generic::IEnumerator<SharedPtr<RevisionGroup>>> e = revisions->get_Groups()->GetEnumerator();
            while (e->MoveNext())
            {
                std::cout << (String::Format(u"\tGroup type \"{0}\", ", e->get_Current()->get_RevisionType()) +
                              String::Format(u"author: {0}, contents: [{1}]", e->get_Current()->get_Author(), e->get_Current()->get_Text().Trim()))
                          << std::endl;
            }
        }

        // Each Run affected by a revision gets its Revision object.
        // The revisions' collection is considerably larger than the condensed form we printed above,
        // depending on how many Runs we have segmented the document into during editing in Microsoft Word.
        ASSERT_EQ(11, revisions->get_Count());
        //ExSkip
        std::cout << "\n" << revisions->get_Count() << " revisions:" << std::endl;

        {
            SharedPtr<System::Collections::Generic::IEnumerator<SharedPtr<Revision>>> e = revisions->GetEnumerator();
            while (e->MoveNext())
            {
                // A StyleDefinitionChange strictly affects styles and not document nodes. This means the ParentStyle
                // attribute will always be in use, while the ParentNode will always be null.
                // Since all other changes affect nodes, ParentNode will conversely be in use, and ParentStyle will be null.
                if (e->get_Current()->get_RevisionType() == RevisionType::StyleDefinitionChange)
                {
                    std::cout << (String::Format(u"\tRevision type \"{0}\", ", e->get_Current()->get_RevisionType()) +
                                  String::Format(u"author: {0}, style: [{1}]", e->get_Current()->get_Author(),
                                                         e->get_Current()->get_ParentStyle()->get_Name()))
                              << std::endl;
                }
                else
                {
                    std::cout << (String::Format(u"\tRevision type \"{0}\", ", e->get_Current()->get_RevisionType()) +
                                  String::Format(u"author: {0}, contents: [{1}]", e->get_Current()->get_Author(),
                                                         e->get_Current()->get_ParentNode()->GetText().Trim()))
                              << std::endl;
                }
            }
        }

        // While the collection of revision groups provides a clearer overview of all revisions that took place in the document,
        // the changes must be accepted/rejected by the revisions themselves, the RevisionCollection, or the document.
        // In this case, we will reject all revisions via the collection, reverting the document to its original form.
        revisions->RejectAll();

        ASSERT_EQ(0, revisions->get_Count());
        //ExEnd
    }

    void GetInfoAboutRevisionsInRevisionGroups()
    {
        //ExStart
        //ExFor:RevisionGroup
        //ExFor:RevisionGroup.Author
        //ExFor:RevisionGroup.RevisionType
        //ExFor:RevisionGroup.Text
        //ExFor:RevisionGroupCollection
        //ExFor:RevisionGroupCollection.Count
        //ExSummary:Shows how to get info about a group of revisions in document.
        auto doc = MakeObject<Document>(MyDir + u"Revisions.docx");

        ASSERT_EQ(7, doc->get_Revisions()->get_Groups()->get_Count());

        for (auto group : System::IterateOver(doc->get_Revisions()->get_Groups()))
        {
            std::cout << String::Format(u"Revision author: {0}; Revision type: {1} \n\tRevision text: {2}", group->get_Author(),
                                                group->get_RevisionType(), group->get_Text())
                      << std::endl;
        }
        //ExEnd
    }

    void GetSpecificRevisionGroup()
    {
        //ExStart
        //ExFor:RevisionGroupCollection
        //ExFor:RevisionGroupCollection.Item(Int32)
        //ExSummary:Shows how to get a group of revisions in document.
        auto doc = MakeObject<Document>(MyDir + u"Revisions.docx");

        SharedPtr<RevisionGroup> revisionGroup = doc->get_Revisions()->get_Groups()->idx_get(0);
        //ExEnd

        ASSERT_EQ(RevisionType::Deletion, revisionGroup->get_RevisionType());
        ASSERT_EQ(u"Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur. ", revisionGroup->get_Text());
    }

    void ShowRevisionBalloons()
    {
        //ExStart
        //ExFor:RevisionOptions.ShowInBalloons
        //ExSummary:Shows how display revisions in balloons.
        auto doc = MakeObject<Document>(MyDir + u"Revisions.docx");

        // By default, revisions are identifiable by different text colors.
        // Set a revision option to show more details about each revision in a balloon on the right margin of the page.
        doc->get_LayoutOptions()->get_RevisionOptions()->set_ShowInBalloons(ShowInBalloons::FormatAndDelete);
        doc->Save(ArtifactsDir + u"Revision.ShowRevisionBalloons.pdf");
        //ExEnd
    }

    void RevisionOptions_()
    {
        //ExStart
        //ExFor:ShowInBalloons
        //ExFor:RevisionOptions.ShowInBalloons
        //ExFor:RevisionOptions.CommentColor
        //ExFor:RevisionOptions.DeletedTextColor
        //ExFor:RevisionOptions.DeletedTextEffect
        //ExFor:RevisionOptions.InsertedTextEffect
        //ExFor:RevisionOptions.MovedFromTextColor
        //ExFor:RevisionOptions.MovedFromTextEffect
        //ExFor:RevisionOptions.MovedToTextColor
        //ExFor:RevisionOptions.MovedToTextEffect
        //ExFor:RevisionOptions.RevisedPropertiesColor
        //ExFor:RevisionOptions.RevisedPropertiesEffect
        //ExFor:RevisionOptions.RevisionBarsColor
        //ExFor:RevisionOptions.RevisionBarsWidth
        //ExFor:RevisionOptions.ShowOriginalRevision
        //ExFor:RevisionOptions.ShowRevisionMarks
        //ExFor:RevisionTextEffect
        //ExSummary:Shows how to edit appearance of revisions.
        auto doc = MakeObject<Document>(MyDir + u"Revisions.docx");

        // Get the RevisionOptions object that controls the appearance of revisions.
        SharedPtr<RevisionOptions> revisionOptions = doc->get_LayoutOptions()->get_RevisionOptions();

        // Render insertion revisions in green and italic.
        revisionOptions->set_InsertedTextColor(RevisionColor::Green);
        revisionOptions->set_InsertedTextEffect(RevisionTextEffect::Italic);

        // Render deletion revisions in red and bold.
        revisionOptions->set_DeletedTextColor(RevisionColor::Red);
        revisionOptions->set_DeletedTextEffect(RevisionTextEffect::Bold);

        // In a movement revision, the same text will appear twice:
        // once at the departure point and once at the arrival destination.
        // Render the text at the moved-from revision yellow with a double strike through
        // and double-underlined blue at the moved-to revision.
        revisionOptions->set_MovedFromTextColor(RevisionColor::Yellow);
        revisionOptions->set_MovedFromTextEffect(RevisionTextEffect::DoubleStrikeThrough);
        revisionOptions->set_MovedToTextColor(RevisionColor::Blue);
        revisionOptions->set_MovedFromTextEffect(RevisionTextEffect::DoubleUnderline);

        // Render format revisions in dark red and bold.
        revisionOptions->set_RevisedPropertiesColor(RevisionColor::DarkRed);
        revisionOptions->set_RevisedPropertiesEffect(RevisionTextEffect::Bold);

        // Place a thick dark blue bar on the left side of the page next to lines affected by revisions.
        revisionOptions->set_RevisionBarsColor(RevisionColor::DarkBlue);
        revisionOptions->set_RevisionBarsWidth(15.0f);

        // Show revision marks and original text.
        revisionOptions->set_ShowOriginalRevision(true);
        revisionOptions->set_ShowRevisionMarks(true);

        // Get movement, deletion, formatting revisions and comments to show up in green balloons
        // on the right side of the page.
        revisionOptions->set_ShowInBalloons(ShowInBalloons::Format);
        revisionOptions->set_CommentColor(RevisionColor::BrightGreen);

        // These features are only applicable to formats such as .pdf or .jpg.
        doc->Save(ArtifactsDir + u"Revision.RevisionOptions.pdf");
        //ExEnd
    }
};

} // namespace ApiExamples
