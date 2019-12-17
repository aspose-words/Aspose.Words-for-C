#include "stdafx.h"
#include "examples.h"

#include <Model/Document/BreakType.h>
#include <Model/Document/Document.h>
#include <Model/Document/DocumentBuilder.h>
#include <Model/Fields/Nodes/FieldSeparator.h>
#include <Model/Fields/Field.h>
#include <Model/Sections/HeaderFooterType.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Fields;

void InsertNestedFields()
{
    std::cout << "InsertNestedFields example started." << std::endl;
    // ExStart:InsertNestedFields
    // The path to the documents directory.
    System::String outputDataDir = GetOutputDataDir_WorkingWithFields();

    System::SharedPtr<Document> doc = System::MakeObject<Document>();
    System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);

    // Insert a few page breaks (just for testing)
    for (int32_t i = 0; i < 5; i++)
    {
        builder->InsertBreak(BreakType::PageBreak);
    }

    // Move the DocumentBuilder cursor into the primary footer.
    builder->MoveToHeaderFooter(HeaderFooterType::FooterPrimary);

    // We want to insert a field like this:
    // { IF {PAGE} <> {NUMPAGES} "See Next Page" "Last Page" }
    System::SharedPtr<Field> field = builder->InsertField(u"IF ");
    builder->MoveTo(field->get_Separator());
    builder->InsertField(u"PAGE");
    builder->Write(u" <> ");
    builder->InsertField(u"NUMPAGES");
    builder->Write(u" \"See Next Page\" \"Last Page\" ");

    // Finally update the outer field to recalculate the final value. Doing this will automatically update
    // The inner fields at the same time.
    field->Update();
    System::String outputPath = outputDataDir + u"InsertNestedFields.docx";
    doc->Save(outputPath);
    // ExEnd:InsertNestedFields
    std::cout << "Inserted nested fields in the document successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    std::cout << "InsertNestedFields example finished." << std::endl << std::endl;
}