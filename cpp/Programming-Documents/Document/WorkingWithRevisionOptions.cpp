#include "stdafx.h"
#include "examples.h"

#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Layout/Public/LayoutOptions.h>
#include <Aspose.Words.Cpp/Layout/Public/RevisionOptions.h>

using namespace Aspose::Words;

namespace
{
    void SetShowInBalloons(System::String const &inputDataDir, System::String const &outputDataDir)
    {
        // ExStart:SetShowInBalloons
        System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"Revisions.docx");

		// Get the RevisionOptions object that controls the appearance of revisions
        auto revisionOptions = doc->get_LayoutOptions()->get_RevisionOptions();

        // Show deletion revisions in balloon
        revisionOptions->set_ShowInBalloons(Layout::ShowInBalloons::FormatAndDelete);

        System::String outputPath = outputDataDir + u"Revisions.ShowRevisionsInBalloons_out.pdf";
        doc->Save(outputPath);
        // ExEnd:SetShowInBalloons
    }

	void SetMeasurementUnit(System::String const &inputDataDir, System::String const &outputDataDir)
	{
		// ExStart:SetMeasurementUnit
        System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"Input.docx");

		// Set Measurement Units to Inches
		doc->get_LayoutOptions()->get_RevisionOptions()->set_MeasurementUnit(MeasurementUnits::Inches);
		// Show deletion revisions in balloon
		doc->get_LayoutOptions()->get_RevisionOptions()->set_ShowInBalloons(Layout::ShowInBalloons::FormatAndDelete);
		// Show Comments
		doc->get_LayoutOptions()->set_ShowComments(true);

		doc->Save(outputDataDir + u"Revisions.SetMeasurementUnit_out.pdf");
		// ExEnd:SetMeasurementUnit
	}

	void SetRevisionBarsPosition(System::String const &inputDataDir, System::String const &outputDataDir)
	{
		// ExStart:SetRevisionBarsPosition
        System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"Input.docx");

		//Renders revision bars on the right side of a page.
		doc->get_LayoutOptions()->get_RevisionOptions()->set_RevisionBarsPosition(Drawing::HorizontalAlignment::Right);

		doc->Save(outputDataDir + u"Revisions.SetRevisionBarsPosition_out.pdf");
		// ExEnd:SetRevisionBarsPosition
	}


}

void WorkingWithRevisionOptions()
{
	std::cout << "WorkingWithRevisionOptions example started.\n";

	// The path to the documents directories.
	System::String inputDataDir = GetInputDataDir_WorkingWithDocument();
	System::String outputDataDir = GetOutputDataDir_WorkingWithDocument();

	SetShowInBalloons(inputDataDir, outputDataDir);
	SetMeasurementUnit(inputDataDir, outputDataDir);
	SetRevisionBarsPosition(inputDataDir, outputDataDir);

	std::cout << "WorkingWithRevisionOptions example finished.\n\n";
}

