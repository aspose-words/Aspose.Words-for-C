#pragma once

#include <system/string.h>
#include <system/shared_ptr.h>
#include <system/object.h>
#include <system/collections/list.h>
#include <system/collections/ilist.h>

namespace Aspose { namespace Words { class Node; } }
namespace Aspose { namespace Words { class Paragraph; } }
namespace Aspose { namespace Words { class Document; } }
namespace Aspose { namespace Words { class CompositeNode; } }

std::vector<System::SharedPtr<Aspose::Words::Node>> ExtractContent(System::SharedPtr<Aspose::Words::Node> startNode, System::SharedPtr<Aspose::Words::Node> endNode, bool isInclusive);
std::vector<System::SharedPtr<Aspose::Words::Paragraph>> ParagraphsByStyleName(const System::SharedPtr<Aspose::Words::Document>& doc, const System::String& styleName);
System::SharedPtr<Aspose::Words::Document> GenerateDocument(const System::SharedPtr<Aspose::Words::Document>& srcDoc, const std::vector<System::SharedPtr<Aspose::Words::Node>>& nodes);


