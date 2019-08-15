#include "stdafx.h"
#include "examples.h"

#include <Model/Revisions/RevisionCollection.h>
#include <Model/Document/Document.h>

using namespace Aspose::Words;

namespace {
	void NormalComparison(const System::String& dataDir)
	{
		std::cout << "NormalComparison example started." << std::endl;
		// ExStart:NormalComparison
		// Load the document from disk.
		System::SharedPtr<Document> docA = System::MakeObject<Document>(dataDir + u"TestFile.doc");
		System::SharedPtr<Document> docB = System::MakeObject<Document>(dataDir + u"TestFile - Copy.doc");

		// DocA now contains changes as revisions. 
		docA->Compare(docB, u"User", System::DateTime::get_Now());
		// ExEnd:NormalComparison
		std::cout << "NormalComparison example finished." << std::endl << std::endl;
	}

	void CompareForEqual(const System::String& dataDir)
	{
		// ExStart:CompareForEqual
		System::SharedPtr<Document> docA = System::MakeObject<Document>(dataDir + u"TestFile.doc");
		System::SharedPtr<Document> docB = System::MakeObject<Document>(dataDir + u"TestFile - Copy.doc");

		// DocA now contains changes as revisions. 
		if (docA->get_Revisions()->get_Count() == 0)
			std::cout << "Documents are equal" << std::endl << std::endl;
		else
			std::cout << "Documents are not equal" << std::endl << std::endl;
		// ExEnd:CompareForEqual                     
	}
}
void CompareDocument()
{
	std::cout << "CompareDocument example started." << std::endl;
	// The path to the documents directory.
	System::String dataDir = GetDataDir_WorkingWithDocument();
	// Normal Comparison
	NormalComparison(dataDir);
	// Compare For Equal
	CompareForEqual(dataDir);
	std::cout << "CompareDocument example finished." << std::endl << std::endl;
}
